"use strict";

const puppeteer = require("puppeteer");

/**
 * Runs INSIDE the page (Chromium) via page.evaluate. Must be fully
 * self-contained (no closures over outer-scope variables) because
 * Puppeteer serializes it to a string and executes it in the browser.
 *
 * Walks the DOM under each "slide root" element and produces a flat,
 * paint-ordered list of primitive drawables: box (rect/roundRect/ellipse),
 * text (with per-run styling for simple inline emphasis), image, svg.
 *
 * This is the generic, template-agnostic half of the pipeline: it makes
 * no assumption about class names, only about visual geometry & computed
 * style, so it works on arbitrary hand-built HTML decks, not just ones
 * that follow our in-house naming convention.
 */
function browserSideExtract(opts) {
  const { slideSelectorCandidates, inlineTags } = opts;

  function pickSlideRoots() {
    for (const sel of slideSelectorCandidates) {
      const found = Array.from(document.querySelectorAll(sel));
      if (found.length) return found;
    }
    // fallback: direct children of body
    return Array.from(document.body.children);
  }

  function isVisible(el, cs) {
    if (cs.display === "none" || cs.visibility === "hidden") return false;
    if (parseFloat(cs.opacity) === 0) return false;
    const r = el.getBoundingClientRect();
    return r.width > 0.5 && r.height > 0.5;
  }

  function relRect(el, originRect) {
    const r = el.getBoundingClientRect();
    return {
      x: r.left - originRect.left,
      y: r.top - originRect.top,
      w: r.width,
      h: r.height,
    };
  }

  function hasOwnVisualStyle(cs) {
    const bg = cs.backgroundColor;
    const bgImg = cs.backgroundImage;
    const hasBg = bg && bg !== "rgba(0, 0, 0, 0)" && bg !== "transparent";
    const hasBgImage = bgImg && bgImg !== "none";
    const hasBorder =
      (parseFloat(cs.borderTopWidth) > 0 && cs.borderTopStyle !== "none") ||
      (parseFloat(cs.borderLeftWidth) > 0 && cs.borderLeftStyle !== "none") ||
      (parseFloat(cs.borderRightWidth) > 0 && cs.borderRightStyle !== "none") ||
      (parseFloat(cs.borderBottomWidth) > 0 && cs.borderBottomStyle !== "none");
    const hasShadow = cs.boxShadow && cs.boxShadow !== "none";
    return { hasBg, hasBgImage, hasBorder, hasShadow, visual: hasBg || hasBgImage || hasBorder || hasShadow };
  }

  function isInlineOnly(el) {
    return Array.from(el.children).every((c) => inlineTags.includes(c.tagName));
  }

  function extractRuns(el) {
    // Best-effort per-run styling: walk direct child nodes, treat text
    // nodes and inline elements (B/STRONG/EM/SPAN/U/SMALL) as runs so a
    // sentence like `foo <b style="color:red">bar</b> baz` keeps its
    // emphasis instead of collapsing to one flat string.
    const runs = [];
    function walk(node, inherited) {
      if (node.nodeType === Node.TEXT_NODE) {
        const t = node.textContent.replace(/\s+/g, " ");
        if (t.trim() !== "" || (t !== "" && runs.length)) {
          runs.push({ text: t, style: inherited });
        }
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return;
      if (node.tagName === "BR") {
        runs.push({ text: "\n", style: inherited, isBreak: true });
        return;
      }
      const cs = getComputedStyle(node);
      const style = {
        color: cs.color,
        bold: parseInt(cs.fontWeight, 10) >= 600,
        italic: cs.fontStyle === "italic",
        underline: cs.textDecorationLine && cs.textDecorationLine.includes("underline"),
      };
      Array.from(node.childNodes).forEach((child) => walk(child, style));
    }
    const rootCs = getComputedStyle(el);
    const rootStyle = {
      color: rootCs.color,
      bold: parseInt(rootCs.fontWeight, 10) >= 600,
      italic: rootCs.fontStyle === "italic",
      underline: rootCs.textDecorationLine && rootCs.textDecorationLine.includes("underline"),
    };
    Array.from(el.childNodes).forEach((child) => walk(child, rootStyle));
    return runs;
  }

  function shapeTypeFor(cs, rect) {
    const brTL = parseFloat(cs.borderTopLeftRadius) || 0;
    const isCircleish =
      cs.borderRadius && brTL >= Math.min(rect.w, rect.h) / 2 - 1 && Math.abs(rect.w - rect.h) < Math.max(rect.w, rect.h) * 0.25;
    if (isCircleish) return "ellipse";
    if (brTL > 1) return "roundRect";
    return "rect";
  }

  function walkSlide(root, out) {
    const originRect = root.getBoundingClientRect();

    function visit(el, depth) {
      const cs = getComputedStyle(el);
      if (!isVisible(el, cs)) return;

      const tag = el.tagName;
      const rect = relRect(el, originRect);

      if (tag === "IMG") {
        out.push({
          type: "image",
          rect,
          src: el.currentSrc || el.src,
          zIndex: parseInt(cs.zIndex, 10) || 0,
          order: out.length,
        });
        return;
      }

      if (tag === "SVG") {
        out.push({
          type: "svg",
          rect,
          outerHTML: el.outerHTML,
          zIndex: parseInt(cs.zIndex, 10) || 0,
          order: out.length,
        });
        return;
      }

      const vis = hasOwnVisualStyle(cs);
      const ownTextEl = isInlineOnly(el) && el.textContent && el.textContent.trim() !== "";

      if (vis.visual) {
        out.push({
          type: "box",
          shape: shapeTypeFor(cs, rect),
          rect,
          style: {
            backgroundColor: vis.hasBg ? cs.backgroundColor : null,
            backgroundImage: vis.hasBgImage ? cs.backgroundImage : null,
            borderColor: cs.borderTopColor,
            borderWidth: parseFloat(cs.borderTopWidth) || 0,
            borderStyle: cs.borderTopStyle,
            borderRadiusPx: parseFloat(cs.borderTopLeftRadius) || 0,
          },
          zIndex: parseInt(cs.zIndex, 10) || 0,
          order: out.length,
        });
      }

      if (ownTextEl) {
        out.push({
          type: "text",
          rect,
          text: el.innerText,
          runs: extractRuns(el),
          style: {
            color: cs.color,
            fontSizePx: parseFloat(cs.fontSize),
            fontWeight: cs.fontWeight,
            fontStyle: cs.fontStyle,
            fontFamily: cs.fontFamily,
            textAlign: cs.textAlign,
            lineHeight: cs.lineHeight,
            letterSpacingPx: cs.letterSpacing === "normal" ? 0 : parseFloat(cs.letterSpacing) || 0,
          },
          zIndex: parseInt(cs.zIndex, 10) || 0,
          order: out.length,
        });
        return; // don't recurse further into pure text leaves
      }

      // container: recurse into children regardless of whether this
      // element itself produced a box, so nested cards/badges/images
      // are still captured.
      Array.from(el.children).forEach((child) => visit(child, depth + 1));
    }

    const rootCs = getComputedStyle(root);
    visit(root, 0);
    return {
      widthPx: root.getBoundingClientRect().width,
      heightPx: root.getBoundingClientRect().height,
      backgroundColor: rootCs.backgroundColor,
      backgroundImage: rootCs.backgroundImage !== "none" ? rootCs.backgroundImage : null,
    };
  }

  const roots = pickSlideRoots();
  return roots.map((root) => {
    const elements = [];
    const meta = walkSlide(root, elements);
    // stable-sort: DOM order already ~= paint order; explicit z-index
    // only reorders elements that set it, ties keep DOM order.
    elements.sort((a, b) => a.zIndex - b.zIndex || a.order - b.order);
    return { ...meta, elements };
  });
}

const DEFAULT_SLIDE_SELECTORS = [
  ".slide",
  "[data-slide]",
  "section.slide",
  "section",
  ".page",
];

const INLINE_TAGS = ["B", "STRONG", "EM", "I", "SPAN", "U", "SMALL", "BR", "MARK"];

/**
 * @param {string} html - full HTML document (or fragment) to render
 * @param {object} [options]
 * @param {string[]} [options.slideSelectors] - CSS selector candidates,
 *   tried in order; first one that matches >=1 element wins. Falls back
 *   to `body > *` if none match, so arbitrary decks without a `.slide`
 *   convention still produce *something*.
 * @param {string} [options.executablePath] - system Chromium path, for
 *   environments where Puppeteer's bundled browser wasn't downloaded
 *   (e.g. offline Docker build using `apt-get install chromium`).
 * @returns {Promise<Array<{widthPx:number, heightPx:number, backgroundColor:string, backgroundImage:string|null, elements: object[]}>>}
 */
async function extractDom(html, options = {}) {
  const browser = await puppeteer.launch({
    headless: true,
    executablePath: options.executablePath || process.env.PUPPETEER_EXECUTABLE_PATH || undefined,
    args: ["--no-sandbox", "--disable-setuid-sandbox", "--force-color-profile=srgb"],
  });
  try {
    const page = await browser.newPage();
    await page.setViewport({ width: 1920, height: 1080, deviceScaleFactor: 1 });
    await page.setContent(html, { waitUntil: "domcontentloaded", timeout: options.timeoutMs || 60000 });
    // let web fonts / late layout settle
    await page.evaluate(() => document.fonts && document.fonts.ready);

    const slides = await page.evaluate(browserSideExtract, {
      slideSelectorCandidates: options.slideSelectors || DEFAULT_SLIDE_SELECTORS,
      inlineTags: INLINE_TAGS,
    });
    return slides;
  } finally {
    await browser.close();
  }
}

module.exports = { extractDom, DEFAULT_SLIDE_SELECTORS };
