"use strict";

const JSZip = require("jszip");

/**
 * pptxgenjs has no public API for gradient fills (verified: no `gradient`
 * option anywhere in its type defs), so `genericGenerator.js` approximates
 * every CSS `linear-gradient(...)` / `radial-gradient(...)` background as a
 * single solid color. That's a reasonable safe default, but it means
 * slide/cover-style hero backgrounds lose their gradient entirely and just
 * look flat.
 *
 * PPTX itself (OOXML) supports real gradient fills (`<a:gradFill>`), so
 * this module patches the already-built .pptx buffer's slide XML directly,
 * replacing the flat `<a:solidFill>` pptxgenjs wrote for `slide.background`
 * with a proper multi-stop `<a:gradFill>` for slides whose source HTML had
 * a real gradient. This only touches whole-slide backgrounds (the common
 * "cover slide hero" case); per-shape background-image gradients remain
 * the solid-color approximation.
 */

/** CSS `linear-gradient(Ndeg, ...)`: 0deg = pointing up, increases
 *  clockwise. OOXML `<a:lin ang="...">`: 0 = pointing right (3 o'clock),
 *  increases clockwise (Y+ is down in drawing space). Converting between
 *  the two conventions: ooxmlDeg = cssDeg - 90 (mod 360). */
function cssAngleToOoxmlUnits(cssAngleDeg) {
  const ooxmlDeg = (((cssAngleDeg - 90) % 360) + 360) % 360;
  return Math.round(ooxmlDeg * 60000);
}

function buildGsLst(stops) {
  return stops
    .map((s) => `<a:gs pos="${Math.max(0, Math.min(100000, Math.round(s.pos * 1000)))}"><a:srgbClr val="${s.color.hex}"/></a:gs>`)
    .join("");
}

function buildGradFillXml(patch) {
  const gsLst = buildGsLst(patch.stops);
  if (patch.isRadial) {
    const cx = patch.centerPct ? patch.centerPct.xPct : 50;
    const cy = patch.centerPct ? patch.centerPct.yPct : 50;
    // OOXML has no direct "radial-gradient(at X% Y%)" primitive; a
    // path="circle" gradient with fillToRect collapsed toward the focal
    // point is the closest native equivalent PowerPoint supports.
    const l = Math.round(Math.max(0, Math.min(100, cx)) * 1000);
    const t = Math.round(Math.max(0, Math.min(100, cy)) * 1000);
    const r = Math.round(Math.max(0, Math.min(100, 100 - cx)) * 1000);
    const b = Math.round(Math.max(0, Math.min(100, 100 - cy)) * 1000);
    return `<a:gradFill rotWithShape="1"><a:gsLst>${gsLst}</a:gsLst><a:path path="circle"><a:fillToRect l="${l}" t="${t}" r="${r}" b="${b}"/></a:path></a:gradFill>`;
  }
  const ang = cssAngleToOoxmlUnits(patch.angleDeg);
  return `<a:gradFill rotWithShape="1"><a:gsLst>${gsLst}</a:gsLst><a:lin ang="${ang}" scaled="1"/></a:gradFill>`;
}

/**
 * @param {Buffer} buffer - a .pptx file as produced by `pptx.write({ outputType: "nodebuffer" })`
 * @param {Array<{slideIndex:number, isRadial:boolean, angleDeg:number, stops:object[], centerPct?:{xPct:number,yPct:number}}>} patches
 * @returns {Promise<Buffer>}
 */
async function applyGradientPatches(buffer, patches) {
  if (!patches || !patches.length) return buffer;
  const zip = await JSZip.loadAsync(buffer);
  let touched = 0;

  for (const patch of patches) {
    if (!patch.stops || patch.stops.length < 2) continue;
    const path = `ppt/slides/slide${patch.slideIndex + 1}.xml`;
    const file = zip.file(path);
    if (!file) continue;

    // eslint-disable-next-line no-await-in-loop
    let xml = await file.async("string");
    const gradXml = buildGradFillXml(patch);

    if (xml.includes("<p:bg>")) {
      const next = xml.replace(
        /(<p:bg><p:bgPr>)(?:<a:solidFill>[\s\S]*?<\/a:solidFill>|<a:blipFill[\s\S]*?<\/a:blipFill>)([\s\S]*?<\/p:bgPr><\/p:bg>)/,
        `$1${gradXml}$2`
      );
      if (next !== xml) {
        xml = next;
        touched++;
      }
    } else {
      // pptxgenjs only emits <p:bg> when a background color/image was set;
      // if there was none at all, insert a fresh one right after <p:cSld>,
      // which the schema requires to come immediately before <p:spTree>.
      const next = xml.replace(/(<p:cSld[^>]*>)/, `$1<p:bg><p:bgPr>${gradXml}<a:effectLst/></p:bgPr></p:bg>`);
      if (next !== xml) {
        xml = next;
        touched++;
      }
    }

    zip.file(path, xml);
  }

  if (!touched) return buffer;
  return zip.generateAsync({ type: "nodebuffer" });
}

module.exports = { applyGradientPatches, cssAngleToOoxmlUnits, buildGradFillXml };
