"use strict";

/**
 * Specialized generator for MaterialPR's in-house "PR event proposal deck"
 * HTML template family (the `.slide` / `.s-head` / `.yoken` / `.ro2-body` /
 * `.pf-body` ... class-name system used across their proposal decks).
 *
 * Unlike the generic engine (server/generator/genericGenerator.js), this
 * module understands the *semantics* of each block type (a Q&A card, a
 * run-of-show timeline row, a quotation table...) because it was written
 * by hand against this specific template - it parses real content out of
 * the supplied HTML with cheerio (no headless browser needed, since we
 * already know the fixed 1920x1080 layout and don't need computed CSS),
 * and reproduces the polished, hand-tuned layouts validated in this
 * project's earlier manual conversion. Content changes (new copy, new
 * numbers, new images) are picked up automatically; layout STRUCTURE
 * changes (a genuinely new slide type) are not - those fall through to
 * the generic per-slide renderer instead of failing outright.
 */

const cheerio = require("cheerio");
const pptxgen = require("pptxgenjs");
const { registerTemplate } = require("./registry");
const { renderSlideModelIntoPptx } = require("../genericGenerator");
const { extractDom } = require("../../extractor/extractDom");

// ---------- fingerprint ----------
const SIGNATURE_TOKENS = [
  "ro2-body", "spk2-body", "pf-idx", "ad-visual", "yoken", "qt-tbl", "cw-tbl", "ai-body",
];
function fingerprint(html) {
  const hits = SIGNATURE_TOKENS.filter((tok) => html.includes(tok));
  return hits.length >= 4; // require a strong match to avoid false positives
}

// ---------- geometry / palette (matches the source deck's own CSS) ----------
const S = 13.333 / 1920;
const IN = (px) => +(px * S).toFixed(3);
const PT = (px) => Math.round(px * 0.5);

const RED = "B50012";
const RED2 = "FF0D14";
const INK = "1A1A1A";
const BODY = "555555";
const BODY2 = "666666";
const MUTED = "999999";
const CARD_BG = "FAF7F7";
const CARD_BORDER = "F0E6E6";
const WHITE = "FFFFFF";
const F_JP = "Yu Gothic";
const F_EN = "Arial";

const ROLE_COLORS = {
  guest: { fill: RED, text: WHITE, line: null },
  host: { fill: "1F2A44", text: WHITE, line: null },
  mc: { fill: WHITE, text: "555555", line: { color: "B8B8B8", width: 1.5 } },
  staff: { fill: "E7E7E7", text: "888888", line: { color: "DCDCDC", width: 1 } },
  talent: { fill: RED, text: WHITE, line: null },
  kol: { fill: "1A7A4A", text: WHITE, line: null },
};

// ---------- cheerio helpers ----------

/** Splits an element's innerHTML on <br> and strips remaining tags per
 *  line - used for headings that only ever contain a manual line break. */
function linesFromBr($, el) {
  const html = $(el).html() || "";
  return html.split(/<br\s*\/?>/i).map((h) => cheerio.load(`<div>${h}</div>`)("div").text().trim());
}

/** Walks an element's direct child nodes and turns them into a pptxgenjs
 *  rich-text run array, treating <b>/<strong> as the template's standard
 *  "bold + accent color" emphasis (the CSS rule `... b{color:#B50012}`
 *  repeated throughout this template family) unless overridden. */
function inlineRuns($, el, opts = {}) {
  const baseColor = opts.color || INK;
  const emphasisColor = opts.emphasisColor || RED;
  const emphasisUnderline = !!opts.emphasisUnderline;
  const runs = [];
  $(el)
    .contents()
    .each((_, node) => {
      if (node.type === "text") {
        const t = node.data.replace(/\s+/g, (m) => (m.includes("\n") ? " " : m));
        if (t !== "") runs.push({ text: t, options: { color: baseColor } });
      } else if (node.type === "tag" && (node.name === "b" || node.name === "strong")) {
        runs.push({
          text: $(node).text(),
          options: { color: emphasisColor, bold: true, underline: emphasisUnderline ? { style: "sng" } : undefined },
        });
      } else if (node.type === "tag" && node.name === "br") {
        if (runs.length) runs[runs.length - 1].options.breakLine = true;
      } else if (node.type === "tag") {
        runs.push({ text: $(node).text(), options: { color: baseColor } });
      }
    });
  return runs.length ? runs : [{ text: $(el).text().trim(), options: { color: baseColor } }];
}

function pill(slide, x, y, w, h, text, opts = {}) {
  slide.addShape("roundRect", {
    x, y, w, h, fill: { color: opts.fill || WHITE }, rectRadius: h / 2,
    line: opts.line || { color: "F0DADA", width: 1.5 },
  });
  slide.addText(text, {
    x, y, w, h, align: "center", valign: "middle",
    fontFace: F_JP, fontSize: opts.fontSize || 10, bold: true, color: opts.color || RED,
  });
}

function header(pptx, slide, eyebrow, titleLines, opts = {}) {
  slide.addText(eyebrow, {
    x: IN(96), y: IN(58), w: IN(1500), h: IN(36),
    fontFace: F_EN, fontSize: (opts.eyebrowSize || 12), bold: true, color: RED2, charSpacing: 1,
  });
  slide.addText(Array.isArray(titleLines) ? titleLines.join("\n") : titleLines, {
    x: IN(96), y: IN(94), w: opts.titleW || IN(1728), h: opts.titleH || IN(120),
    fontFace: F_JP, fontSize: opts.titleSize || PT(48), bold: true, color: INK,
    lineSpacingMultiple: 1.15, valign: "top",
  });
  slide.addShape("roundRect", {
    x: IN(96), y: opts.ruleY || IN(214), w: IN(74), h: IN(7),
    fill: { color: RED2 }, line: { type: "none" }, rectRadius: 0.04,
  });
}

function footer(slide, pageNo, total) {
  slide.addShape("line", {
    x: IN(96), y: IN(1010), w: IN(1728), h: 0, line: { color: "E6E6E6", width: 1 },
  });
  slide.addText(`${pageNo} / ${total}`, {
    x: IN(1500), y: IN(1018), w: IN(324), h: IN(30),
    fontFace: F_EN, fontSize: 11, color: MUTED, align: "right",
  });
}

function sectionSlide(pptx, no, en, jpLines) {
  const slide = pptx.addSlide();
  slide.background = { color: RED };
  slide.addText(no, { x: 0.85, y: 1.1, w: 4, h: 1.6, fontFace: F_EN, fontSize: 96, bold: true, color: WHITE, transparency: 72 });
  slide.addText(en, { x: 0.9, y: 2.75, w: 8, h: 0.4, fontFace: F_EN, fontSize: 15, bold: true, color: "FFD9DB", charSpacing: 3 });
  slide.addText(jpLines.join("\n"), { x: 0.9, y: 3.15, w: 9, h: 1.6, fontFace: F_JP, fontSize: 41, bold: true, color: WHITE, lineSpacingMultiple: 1.15 });
  slide.addShape("roundRect", { x: 0.9, y: 4.85, w: 0.83, h: 0.055, fill: { color: WHITE }, line: { type: "none" }, rectRadius: 0.02 });
  return slide;
}

// ---------- per-slide-type generators ----------

function genCover($, slideEl, pptx) {
  const slide = pptx.addSlide();
  slide.background = { color: RED };
  const eyebrow = $(slideEl).find(".cover-eyebrow").first().text().trim();
  const titleLines = linesFromBr($, $(slideEl).find(".cover-title").get(0));
  const ccLbl = $(slideEl).find(".cc-lbl").first().text().trim();
  const ccNm = $(slideEl).find(".cc-nm").first().text().trim();
  const meta = $(slideEl).find(".cover-meta").first().text().trim();
  const logo = $(slideEl).find(".cover-logo").first().text().trim();

  slide.addText(eyebrow, { x: 0.9, y: 1.55, w: 11.8, h: 0.4, fontFace: F_EN, fontSize: 14, bold: true, color: "FFD9DB", charSpacing: 1, wrap: false });
  slide.addText(titleLines.join("\n"), { x: 0.9, y: 2.0, w: 11, h: 1.9, fontFace: F_JP, fontSize: 40, bold: true, color: WHITE, lineSpacingMultiple: 1.2 });
  slide.addShape("roundRect", { x: 0.9, y: 4.05, w: 0.7, h: 0.055, fill: { color: WHITE }, line: { type: "none" }, rectRadius: 0.02 });
  slide.addText([
    { text: `${ccLbl}   `, options: { fontFace: F_EN, fontSize: 12, bold: true, color: "FFD9DB", charSpacing: 2 } },
    { text: ccNm, options: { fontFace: F_JP, fontSize: 16, bold: true, color: WHITE } },
  ], { x: 0.9, y: 4.35, w: 10, h: 0.4 });
  slide.addText(meta, { x: 0.9, y: 6.55, w: 4, h: 0.35, fontFace: F_EN, fontSize: 16, color: "FFD9DB", charSpacing: 1 });
  slide.addText(logo, { x: 9.3, y: 6.55, w: 3.1, h: 0.35, align: "right", fontFace: F_EN, fontSize: 18, bold: true, color: WHITE, charSpacing: 2 });
}

function genSection($, slideEl, pptx) {
  const no = $(slideEl).find(".section-no").first().text().trim();
  const en = $(slideEl).find(".section-en").first().text().trim();
  const jpLines = linesFromBr($, $(slideEl).find(".section-jp").get(0));
  sectionSlide(pptx, no, en, jpLines);
}

function genYoken($, slideEl, pptx, ctx) {
  const slide = pptx.addSlide();
  const eyebrow = $(slideEl).find(".s-eyebrow").first().text().trim();
  const title = $(slideEl).find(".s-title").first().text().trim();
  header(pptx, slide, eyebrow, [title]);

  const rows = $(slideEl).find(".yrow").toArray();
  const top = 2.05, left = 0.667, w = 12, rowH = Math.min(0.85, (6.8 - top) / rows.length);
  slide.addShape("roundRect", { x: left, y: top, w, h: rowH * rows.length, fill: { color: WHITE }, line: { color: "EEEEEE", width: 1 }, rectRadius: 0.06 });
  rows.forEach((row, i) => {
    const y = top + rowH * i;
    if (i > 0) slide.addShape("line", { x: left, y, w, h: 0, line: { color: "EEEEEE", width: 1 } });
    slide.addShape("rect", { x: left, y, w: 1.9, h: rowH, fill: { color: "FBF4F5" }, line: { type: "none" } });
    slide.addText($(row).find(".yk").text().trim(), { x: left + 0.2, y, w: 1.7, h: rowH, valign: "middle", fontFace: F_JP, fontSize: 13, bold: true, color: RED });
    const runs = inlineRuns($, $(row).find(".yv").get(0), { color: INK }).map((r) => ({ text: r.text, options: { ...r.options, fontFace: F_JP, fontSize: 13 } }));
    slide.addText(runs, { x: left + 2.15, y, w: w - 2.35, h: rowH, valign: "middle", lineSpacingMultiple: 1.15 });
  });
  footer(slide, ctx.pageNo(), ctx.total);
}

function genCmp2($, slideEl, pptx, ctx) {
  const slide = pptx.addSlide();
  header(pptx, slide, $(slideEl).find(".s-eyebrow").text().trim(), [$(slideEl).find(".s-title").text().trim()], { titleSize: PT(46) });

  const cols = $(slideEl).find(".cmp2-col").toArray();
  const colY = 2.35, colH = 3.3, colW = 5.7, gap = 0.42, leftX = 0.667;
  cols.forEach((col, i) => {
    const x = leftX + i * (colW + gap);
    slide.addShape("roundRect", { x, y: colY, w: colW, h: colH, fill: { color: CARD_BG }, line: { color: CARD_BORDER, width: 1 }, rectRadius: 0.08 });
    pill(slide, x + 0.4, colY + 0.4, 1.3, 0.4, $(col).find(".tag").text().trim(), { fill: RED, color: WHITE, line: { type: "none" }, fontSize: 11 });
    slide.addText($(col).find(".lb").text().trim(), { x: x + 0.4, y: colY + 0.95, w: colW - 0.8, h: 0.7, fontFace: F_JP, fontSize: 30, bold: true, color: RED });
    slide.addText($(col).find("h4").text().trim(), { x: x + 0.4, y: colY + 1.65, w: colW - 0.8, h: 0.5, fontFace: F_JP, fontSize: 20, bold: true, color: INK });
    slide.addText($(col).find("p").text().trim(), { x: x + 0.4, y: colY + 2.2, w: colW - 0.8, h: 0.9, fontFace: F_JP, fontSize: 13, color: BODY2, lineSpacingMultiple: 1.3 });
  });

  const goalY = colY + colH + 0.35;
  slide.addShape("roundRect", { x: leftX, y: goalY, w: colW * 2 + gap, h: 0.95, fill: { color: RED }, line: { type: "none" }, rectRadius: 0.08 });
  pill(slide, leftX + 0.4, goalY + 0.28, 1.1, 0.4, $(slideEl).find(".cmp2-goal .gl").text().trim(), { fill: WHITE, color: RED, line: { type: "none" }, fontSize: 12 });
  slide.addText($(slideEl).find(".cmp2-goal p").text().trim(), { x: leftX + 1.75, y: goalY, w: colW * 2 + gap - 2.1, h: 0.95, valign: "middle", fontFace: F_JP, fontSize: 15, bold: true, color: WHITE });
  footer(slide, ctx.pageNo(), ctx.total);
}

function genRoleList($, slideEl, pptx, ctx) {
  const slide = pptx.addSlide();
  header(pptx, slide, $(slideEl).find(".s-eyebrow").text().trim(), [$(slideEl).find(".s-title").text().trim()]);
  const rows = $(slideEl).find(".rl-row").toArray();
  const top = 2.15, x = 0.667, w = 12, rowH = Math.min(1.28, (6.9 - top) / rows.length - 0.14);
  rows.forEach((row, i) => {
    const y = top + i * (rowH + 0.14);
    slide.addShape("roundRect", { x, y, w, h: rowH, fill: { color: CARD_BG }, line: { color: CARD_BORDER, width: 1 }, rectRadius: 0.07 });
    slide.addShape("roundRect", { x: x + 0.3, y: y + rowH / 2 - 0.42, w: 0.84, h: 0.84, fill: { color: RED }, line: { type: "none" }, rectRadius: 0.1 });
    slide.addText($(row).find(".rl-no").text().trim(), { x: x + 0.3, y: y + rowH / 2 - 0.42, w: 0.84, h: 0.84, align: "center", valign: "middle", fontFace: F_EN, fontSize: 20, bold: true, color: WHITE });
    slide.addText($(row).find(".rl-tx h4").text().trim(), { x: x + 1.4, y: y + 0.18, w: w - 1.8, h: 0.4, fontFace: F_JP, fontSize: 16, bold: true, color: INK });
    slide.addText($(row).find(".rl-tx p").text().trim(), { x: x + 1.4, y: y + 0.58, w: w - 1.8, h: 0.6, fontFace: F_JP, fontSize: 12, color: BODY2, lineSpacingMultiple: 1.3 });
  });
  footer(slide, ctx.pageNo(), ctx.total);
}

function genSpeakers($, slideEl, pptx, ctx) {
  const slide = pptx.addSlide();
  header(pptx, slide, $(slideEl).find(".s-eyebrow").text().trim(), [$(slideEl).find(".s-title").text().trim()]);
  const rows = $(slideEl).find(".spk2-row").toArray();
  const top = 2.15, x = 0.667, w = 12, rowH = Math.min(1.28, (6.9 - top) / rows.length - 0.14);
  rows.forEach((row, i) => {
    const y = top + i * (rowH + 0.14);
    slide.addShape("roundRect", { x, y, w, h: rowH, fill: { color: CARD_BG }, line: { color: CARD_BORDER, width: 1 }, rectRadius: 0.07 });
    pill(slide, x + 0.3, y + rowH / 2 - 0.24, 1.35, 0.48, $(row).find(".spk2-badge").text().trim(), { fill: RED, color: WHITE, line: { type: "none" }, fontSize: 12 });
    slide.addText([
      { text: `${$(row).find(".spk2-nm").text().trim()}   `, options: { fontFace: F_JP, fontSize: 16, bold: true, color: INK } },
      { text: $(row).find(".spk2-role").text().trim(), options: { fontFace: F_JP, fontSize: 11, color: MUTED } },
    ], { x: x + 1.85, y: y + 0.16, w: w - 4.0, h: 0.35 });
    slide.addText($(row).find(".spk2-desc").text().trim(), { x: x + 1.85, y: y + 0.55, w: w - 4.0, h: 0.6, fontFace: F_JP, fontSize: 12, color: BODY2, lineSpacingMultiple: 1.3 });
    pill(slide, x + w - 1.9, y + rowH / 2 - 0.22, 1.6, 0.44, $(row).find(".spk2-pill").text().trim(), { fill: WHITE, color: RED, line: { color: "F0DADA", width: 1.5 }, fontSize: 11 });
  });
  footer(slide, ctx.pageNo(), ctx.total);
}

function genRunOfShow($, slideEl, pptx, ctx) {
  const slide = pptx.addSlide();
  header(pptx, slide, $(slideEl).find(".s-eyebrow").text().trim(), [$(slideEl).find(".s-title").text().trim()], { titleH: IN(90), ruleY: IN(180) });

  let lx = 0.667;
  $(slideEl).find(".ro2-legend .ro2-lg").each((_, lg) => {
    const swClass = ($(lg).find(".ro2-sw").attr("class") || "").split(/\s+/).find((c) => ROLE_COLORS[c]);
    const c = (ROLE_COLORS[swClass] || {}).fill || "CCCCCC";
    slide.addShape("roundRect", { x: lx, y: 1.62, w: 0.16, h: 0.16, fill: { color: c }, line: { type: "none" }, rectRadius: 0.03 });
    slide.addText($(lg).text().trim(), { x: lx + 0.22, y: 1.56, w: 0.9, h: 0.28, fontFace: F_JP, fontSize: 10, color: BODY });
    lx += 1.05;
  });

  const bands = $(slideEl).find(".ro2-band").toArray();
  const bandHeights = bands.map((b) => 0.6 + $(b).find(".ro2-r").length * 0.42);
  const totalBandH = bandHeights.reduce((a, b) => a + b, 0);
  const availableH = 6.9 - 1.98;
  const scaleH = totalBandH > availableH ? availableH / totalBandH : 1;

  let y0 = 1.98;
  bands.forEach((band, bi) => {
    const h0 = bandHeights[bi] * scaleH;
    slide.addText([
      { text: `${$(band).find(".ro2-b").text().trim()}  `, options: { fontFace: F_EN, fontSize: 11, bold: true, color: RED } },
      { text: $(band).find(".ro2-tt").text().trim(), options: { fontFace: F_JP, fontSize: 13, bold: true, color: INK } },
      { text: `   ${$(band).find(".ro2-sub").text().trim()}`, options: { fontFace: F_JP, fontSize: 10, color: MUTED } },
    ], { x: 0.667, y: y0, w: 12, h: 0.26 });

    const rows = $(band).find(".ro2-r").toArray().map((r) => ([
      { text: $(r).find(".ro2-tm").text().trim().replace(/(\d+)(\D+)/, "$1 $2"), options: { fontFace: F_EN, fontSize: 10, bold: true, color: RED, align: "center", valign: "middle" } },
      { text: $(r).find(".ro2-h").text().trim(), options: { fontFace: F_JP, fontSize: 11, bold: true, color: INK, valign: "middle" } },
      { text: $(r).find(".ro2-d").text().trim(), options: { fontFace: F_JP, fontSize: 9.5, color: MUTED, valign: "middle" } },
      { text: $(r).find(".ro2-chip").map((_, c) => $(c).text().trim()).get().join("・"), options: { fontFace: F_JP, fontSize: 9, bold: true, color: RED, align: "center", valign: "middle" } },
    ]));
    slide.addTable(rows, {
      x: 0.667, y: y0 + 0.3, w: 12, h: h0 - 0.3,
      colW: [0.9, 2.6, 6.4, 2.1],
      border: { type: "solid", color: "EEEEEE", pt: 0.75 },
      fill: { color: "FCFAFA" },
      autoPage: false,
      valign: "middle",
      margin: [3, 6, 3, 6],
    });
    y0 += h0;
  });
  footer(slide, ctx.pageNo(), ctx.total);
}

function genTalkQa($, slideEl, pptx, ctx) {
  const slide = pptx.addSlide();
  header(pptx, slide, $(slideEl).find(".s-eyebrow").text().trim(), [$(slideEl).find(".s-title").text().trim()], { titleSize: PT(46) });
  const leadRuns = inlineRuns($, $(slideEl).find(".tk-lead").get(0), { color: BODY }).map((r) => ({ text: r.text, options: { ...r.options, fontFace: F_JP, fontSize: 13.5 } }));
  slide.addText(leadRuns, { x: 0.667, y: 2.0, w: 12, h: 0.5, lineSpacingMultiple: 1.3 });

  const items = $(slideEl).find(".tk-qa").toArray().map((qa) => ({
    q: $(qa).find(".tk-q").text().trim(),
    a: $(qa).find(".tk-a").toArray().map((a) => ({ who: $(a).find(".tk-who").text().trim(), tx: $(a).find(".tk-atx").text().trim() })),
  }));

  const areaTop = 2.55, areaBottom = 6.95, w = 12, x = 0.667;
  let gap = 0.18, qH = 0.48, aLineH = 0.5;
  let heights = items.map((it) => qH + it.a.length * aLineH);
  let totalH = heights.reduce((a, b) => a + b, 0) + gap * (items.length - 1);
  const available = areaBottom - areaTop;
  if (totalH > available) {
    const scale = available / totalH;
    qH *= scale; aLineH *= scale; gap *= scale;
    heights = items.map((it) => qH + it.a.length * aLineH);
    totalH = heights.reduce((a, b) => a + b, 0) + gap * (items.length - 1);
  }
  let y = areaTop + Math.max(0, (available - totalH) / 2);

  items.forEach((it, i) => {
    const rowH = heights[i];
    slide.addShape("roundRect", { x, y, w, h: rowH, fill: { color: CARD_BG }, line: { color: CARD_BORDER, width: 1 }, rectRadius: 0.06 });
    pill(slide, x + 0.25, y + 0.15, 0.4, 0.4, "Q", { fill: RED, color: WHITE, line: { type: "none" }, fontSize: 13 });
    slide.addText(it.q, { x: x + 0.8, y: y + 0.06, w: w - 1.1, h: qH, fontFace: F_JP, fontSize: 13.5, bold: true, color: INK, valign: "middle" });
    it.a.forEach((a, j) => {
      const ay = y + qH + j * aLineH;
      const pw = Math.max(0.6, 0.135 * a.who.length + 0.32);
      const isTalent = a.who === "モナキ" || a.who.length <= 3 && !/^設問|^MC$/.test(a.who) && a.who !== "MC";
      const usePrimary = a.who !== "MC" && !/^設問/.test(a.who);
      pill(slide, x + 0.55, ay + (aLineH - 0.32) / 2, pw, 0.32, a.who,
        { fill: usePrimary ? RED : WHITE, color: usePrimary ? WHITE : RED, line: usePrimary ? { type: "none" } : { color: "F0DADA", width: 1.25 }, fontSize: 9.5 });
      slide.addText(a.tx, { x: x + 1.65, y: ay, w: w - 1.95, h: aLineH, valign: "middle", fontFace: F_JP, fontSize: 12, color: BODY2, lineSpacingMultiple: 1.2 });
    });
    y += rowH + gap;
  });
  footer(slide, ctx.pageNo(), ctx.total);
}

function genAttractionCards($, slideEl, pptx, ctx) {
  const slide = pptx.addSlide();
  header(pptx, slide, $(slideEl).find(".s-eyebrow").text().trim(), [$(slideEl).find(".s-title").text().trim()], { titleSize: PT(44) });
  const leadRuns = inlineRuns($, $(slideEl).find(".ai-lead").get(0), { color: BODY }).map((r) => ({ text: r.text, options: { ...r.options, fontFace: F_JP, fontSize: 13 } }));
  slide.addText(leadRuns, { x: 0.667, y: 1.95, w: 12, h: 0.55, lineSpacingMultiple: 1.3 });

  const cards = $(slideEl).find(".ai-card").toArray();
  const top = 2.65, cw = (12 - 0.22 * (cards.length - 1)) / cards.length, gap = 0.22, x0 = 0.667, ch = 4.1;
  cards.forEach((card, i) => {
    const x = x0 + i * (cw + gap);
    const imgSrc = $(card).find(".ai-ph img").first().attr("src");
    slide.addShape("roundRect", { x, y: top, w: cw, h: ch, fill: { color: CARD_BG }, line: { color: CARD_BORDER, width: 1 }, rectRadius: 0.08 });
    pill(slide, x + 0.25, top + 0.25, 0.42, 0.42, $(card).find(".ai-no").text().trim(), { fill: RED, color: WHITE, line: { type: "none" }, fontSize: 15 });
    const htLines = linesFromBr($, $(card).find(".ai-ht").get(0));
    slide.addText(htLines.join("\n"), { x: x + 0.8, y: top + 0.2, w: cw - 1.05, h: 0.55, fontFace: F_JP, fontSize: 12.5, bold: true, color: INK, valign: "middle", lineSpacingMultiple: 1.15 });
    const imgH = (cw - 0.5) * 9 / 16;
    if (imgSrc) slide.addImage({ data: imgSrc, x: x + 0.25, y: top + 0.85, w: cw - 0.5, h: imgH, sizing: { type: "cover", w: cw - 0.5, h: imgH } });
    const belowImgY = top + 0.85 + imgH + 0.16;
    pill(slide, x + 0.25, belowImgY, 1.55, 0.34, $(card).find(".ai-tag").text().trim(), { fill: WHITE, color: RED, line: { color: "F0DADA", width: 1.25 }, fontSize: 10.5 });
    slide.addText($(card).find("p").last().text().trim(), { x: x + 0.25, y: belowImgY + 0.42, w: cw - 0.5, h: ch - (belowImgY + 0.42 - top) - 0.15, fontFace: F_JP, fontSize: 11.5, color: BODY2, lineSpacingMultiple: 1.3 });
  });
  footer(slide, ctx.pageNo(), ctx.total);
}

function genAttractionDetail($, slideEl, pptx, ctx) {
  const slide = pptx.addSlide();
  header(pptx, slide, $(slideEl).find(".s-eyebrow").text().trim(), [$(slideEl).find(".s-title").text().trim()], { titleSize: PT(40) });

  const imgSrc = $(slideEl).find(".ad-visual img").first().attr("src");
  const vno = $(slideEl).find(".ad-vno").text().trim();
  const imgX = 0.667, imgY = 1.95, imgW = 5.6, imgH = 3.55;
  if (imgSrc) slide.addImage({ data: imgSrc, x: imgX, y: imgY, w: imgW, h: imgH, sizing: { type: "cover", w: imgW, h: imgH } });
  slide.addShape("roundRect", { x: imgX + 0.18, y: imgY + imgH - 0.55, w: 1.7, h: 0.4, fill: { color: RED }, line: { type: "none" }, rectRadius: 0.05 });
  slide.addText(vno, { x: imgX + 0.18, y: imgY + imgH - 0.55, w: 1.7, h: 0.4, align: "center", valign: "middle", fontFace: F_JP, fontSize: 12, bold: true, color: WHITE });

  const rightX = imgX + imgW + 0.4, rightW = 0.667 + 12 - rightX;
  const steps = $(slideEl).find(".ad-step").toArray();
  const stepTop = 1.95, stepH = 0.72, stepGap = 0.1;
  steps.forEach((s, i) => {
    const y = stepTop + i * (stepH + stepGap);
    slide.addShape("roundRect", { x: rightX, y, w: 0.5, h: 0.5, fill: { color: RED }, line: { type: "none" }, rectRadius: 0.08 });
    slide.addText(String(i + 1), { x: rightX, y, w: 0.5, h: 0.5, align: "center", valign: "middle", fontFace: F_EN, fontSize: 16, bold: true, color: WHITE });
    slide.addText($(s).find("h5").text().trim(), { x: rightX + 0.65, y: y - 0.02, w: rightW - 0.65, h: 0.3, fontFace: F_JP, fontSize: 13.5, bold: true, color: INK });
    slide.addText($(s).find("p").text().trim(), { x: rightX + 0.65, y: y + 0.28, w: rightW - 0.65, h: 0.4, fontFace: F_JP, fontSize: 11, color: BODY2, lineSpacingMultiple: 1.2 });
  });
  const aimY = stepTop + steps.length * (stepH + stepGap) + 0.06;
  slide.addShape("roundRect", { x: rightX, y: aimY, w: rightW, h: 0.85, fill: { color: "FBF4F5" }, line: { color: "F0E0E2", width: 1 }, rectRadius: 0.07 });
  slide.addText($(slideEl).find(".ad-aim .lab").text().trim(), { x: rightX + 0.22, y: aimY + 0.1, w: rightW - 0.4, h: 0.24, fontFace: F_JP, fontSize: 10, bold: true, color: RED });
  slide.addText($(slideEl).find(".ad-aim p").text().trim(), { x: rightX + 0.22, y: aimY + 0.36, w: rightW - 0.4, h: 0.45, fontFace: F_JP, fontSize: 11.5, color: BODY2, lineSpacingMultiple: 1.2 });

  const botY = imgY + imgH + 0.25, botH = 6.95 - botY;
  slide.addShape("roundRect", { x: 0.667, y: botY, w: 8.6, h: botH, fill: { color: CARD_BG }, line: { color: CARD_BORDER, width: 1 }, rectRadius: 0.07 });
  const summaryRuns = inlineRuns($, $(slideEl).find(".ad-scopy").get(0), { color: BODY2, emphasisColor: RED }).map((r) => ({ text: r.text, options: { ...r.options, fontFace: F_JP, fontSize: 10.5, bold: r.options.bold, lineSpacingMultiple: 1.3 } }));
  slide.addText(summaryRuns, { x: 0.667 + 0.3, y: botY + 0.16, w: 8.0, h: botH - 0.3, valign: "middle" });

  const shotX = 0.667 + 8.6 + 0.25, shotW = 0.667 + 12 - shotX;
  slide.addShape("roundRect", { x: shotX, y: botY, w: shotW, h: botH, fill: { color: RED }, line: { type: "none" }, rectRadius: 0.07 });
  slide.addText("SHOT", { x: shotX + 0.25, y: botY + 0.16, w: shotW - 0.5, h: 0.26, fontFace: F_EN, fontSize: 11, bold: true, color: "FFD9DB", charSpacing: 1.5 });
  slide.addText($(slideEl).find(".ad-shot p").text().trim(), { x: shotX + 0.25, y: botY + 0.46, w: shotW - 0.5, h: botH - 0.6, fontFace: F_JP, fontSize: 12, bold: true, color: WHITE, lineSpacingMultiple: 1.25 });
  footer(slide, ctx.pageNo(), ctx.total);
}

function genPromoteFlow($, slideEl, pptx, ctx) {
  const slide = pptx.addSlide();
  header(pptx, slide, $(slideEl).find(".s-eyebrow").text().trim(), [$(slideEl).find(".s-title").text().trim()], { titleSize: PT(40) });

  const leftX = 0.667, leftW = 3.05, top = 2.05, h = 4.9;
  slide.addShape("roundRect", { x: leftX, y: top, w: leftW, h, fill: { color: RED }, line: { type: "none" }, rectRadius: 0.1 });
  slide.addText($(slideEl).find(".pf-left .lb").text().trim(), { x: leftX + 0.35, y: top + 0.4, w: leftW - 0.7, h: 0.3, fontFace: F_EN, fontSize: 11, bold: true, color: "FFD9DB", charSpacing: 1.8 });
  const bgRuns = inlineRuns($, $(slideEl).find(".pf-left .bg").get(0), { color: WHITE, emphasisColor: WHITE, emphasisUnderline: true })
    .map((r) => ({ text: r.text, options: { ...r.options, fontFace: F_JP, fontSize: 19, bold: true } }));
  slide.addText(bgRuns, { x: leftX + 0.35, y: top + 0.85, w: leftW - 0.7, h: 1.9, lineSpacingMultiple: 1.4 });
  slide.addShape("line", { x: leftX + 0.35, y: top + 3.0, w: leftW - 0.7, h: 0, line: { color: "E0748A", width: 0.75 } });
  slide.addText($(slideEl).find(".pf-left .sub").text().trim(), { x: leftX + 0.35, y: top + 3.15, w: leftW - 0.7, h: 1.5, fontFace: F_JP, fontSize: 12.5, color: "FFE3E4", lineSpacingMultiple: 1.5 });

  const steps = $(slideEl).find(".pf-step").toArray();
  const rowX = leftX + leftW + 0.35, rowW = 0.667 + 12 - rowX;
  const rowGap = 0.2, rowH = (h - rowGap * (steps.length - 1)) / steps.length, idxW = 1.05;
  steps.forEach((step, i) => {
    const y = top + i * (rowH + rowGap);
    slide.addShape("roundRect", { x: rowX, y, w: rowW, h: rowH, fill: { color: CARD_BG }, line: { color: CARD_BORDER, width: 1 }, rectRadius: 0.07 });
    slide.addShape("line", { x: rowX + idxW, y: y + 0.16, w: 0, h: rowH - 0.32, line: { color: "F2D6D9", width: 2 } });
    slide.addText($(step).find(".pf-idx .step").text().trim(), { x: rowX, y: y + rowH / 2 - 0.5, w: idxW - 0.1, h: 0.24, align: "center", fontFace: F_EN, fontSize: 9.5, color: RED, charSpacing: 0.5 });
    slide.addText(String(i + 1).padStart(2, "0"), { x: rowX, y: y + rowH / 2 - 0.34, w: idxW - 0.1, h: 0.5, align: "center", fontFace: F_EN, fontSize: 26, bold: true, color: RED });
    slide.addText($(step).find(".pf-idx .en").text().trim(), { x: rowX, y: y + rowH / 2 + 0.16, w: idxW - 0.1, h: 0.22, align: "center", fontFace: F_EN, fontSize: 8.5, color: "C98A90", charSpacing: 0.8 });

    const midX = rowX + idxW + 0.25, midW = rowW - idxW - 0.45;
    slide.addText($(step).find(".pf-h").text().trim(), { x: midX, y: y + 0.14, w: midW, h: 0.34, fontFace: F_JP, fontSize: 15, bold: true, color: INK });
    slide.addShape("line", { x: midX, y: y + 0.52, w: midW, h: 0, line: { color: "DDDDDD", width: 0.75, dashType: "dash" } });

    const tasks = $(step).find(".pf-task").toArray();
    const taskTop = y + 0.66, taskColW = (midW - 0.3) / 2, taskRowH = (rowH - 0.8) / Math.ceil(tasks.length / 2);
    tasks.forEach((task, j) => {
      const col = j % 2, r = Math.floor(j / 2);
      const tx = midX + col * (taskColW + 0.3), ty = taskTop + r * taskRowH;
      slide.addShape("ellipse", { x: tx, y: ty + 0.07, w: 0.07, h: 0.07, fill: { color: RED2 }, line: { type: "none" } });
      slide.addText($(task).find(".tt").text().trim(), { x: tx + 0.16, y: ty - 0.02, w: taskColW - 0.16, h: 0.24, fontFace: F_JP, fontSize: 11, bold: true, color: INK });
      slide.addText($(task).find(".td").text().trim(), { x: tx + 0.16, y: ty + 0.21, w: taskColW - 0.16, h: taskRowH - 0.21, fontFace: F_JP, fontSize: 9.5, color: MUTED, lineSpacingMultiple: 1.25 });
    });
  });
  footer(slide, ctx.pageNo(), ctx.total);
}

function genSchedule($, slideEl, pptx, ctx) {
  const slide = pptx.addSlide();
  header(pptx, slide, $(slideEl).find(".s-eyebrow").text().trim(), [$(slideEl).find(".s-title").text().trim()]);
  const leadRuns = inlineRuns($, $(slideEl).find(".sc-lead").get(0), { color: BODY }).map((r) => ({ text: r.text, options: { ...r.options, fontFace: F_JP, fontSize: 13.5 } }));
  slide.addText(leadRuns, { x: 0.667, y: 2.0, w: 12, h: 0.4 });

  const top = 2.5, x = 0.667, w = 12;
  const head = ["TIMING", "PHASE", "DETAIL"].map((t) => ({ text: t, options: { fontFace: F_EN, fontSize: 11, bold: true, color: MUTED } }));
  slide.addTable([head], { x, y: top, w, h: 0.4, colW: [2.4, 3.6, 6.0], border: { pt: 0, color: WHITE }, fill: { color: WHITE }, valign: "middle", margin: [4, 8, 4, 8] });
  slide.addShape("line", { x, y: top + 0.4, w, h: 0, line: { color: INK, width: 2 } });

  const rows = $(slideEl).find(".sc-tbl tbody tr").toArray();
  const rowH = 0.6;
  rows.forEach((row, i) => {
    const y = top + 0.4 + i * rowH;
    const isKey = ($(row).attr("class") || "").includes("key");
    slide.addShape("rect", { x, y, w, h: rowH, fill: { color: isKey ? RED : (i % 2 ? "FCFAFA" : WHITE) }, line: { type: "none" } });
    slide.addShape("line", { x, y: y + rowH, w, h: 0, line: { color: "EEEEEE", width: 1 } });
    slide.addText($(row).find(".sc-mo").text().trim(), { x: x + 0.15, y, w: 2.2, h: rowH, valign: "middle", fontFace: F_EN, fontSize: 13, bold: true, color: isKey ? WHITE : RED });
    slide.addText($(row).find(".sc-lb").text().trim(), { x: x + 2.5, y, w: 3.4, h: rowH, valign: "middle", fontFace: F_JP, fontSize: 13, bold: true, color: isKey ? WHITE : INK });
    slide.addText($(row).find(".sc-ds").text().trim(), { x: x + 6.1, y, w: 5.8, h: rowH, valign: "middle", fontFace: F_JP, fontSize: 11.5, color: isKey ? "FFE3E4" : MUTED });
  });

  const noteY = top + 0.4 + rows.length * rowH + 0.22;
  slide.addShape("roundRect", { x, y: noteY, w, h: 0.5, fill: { color: "FBF4F5" }, line: { type: "none" }, rectRadius: 0.06 });
  const noteRuns = inlineRuns($, $(slideEl).find(".sc-deadline").get(0), { color: BODY }).map((r) => ({ text: r.text, options: { ...r.options, fontFace: F_JP, fontSize: 11.5 } }));
  slide.addText(noteRuns, { x: x + 0.3, y: noteY, w: w - 0.6, h: 0.5, valign: "middle" });
  footer(slide, ctx.pageNo(), ctx.total);
}

function genQuote($, slideEl, pptx, ctx) {
  const slide = pptx.addSlide();
  header(pptx, slide, $(slideEl).find(".s-eyebrow").text().trim(), [$(slideEl).find(".s-title").text().trim()]);
  slide.addText($(slideEl).find(".lead").first().text().trim(), { x: 0.667, y: 2.0, w: 12, h: 0.4, fontFace: F_JP, fontSize: 13.5, color: BODY });

  const top = 2.6, x = 0.667, w = 12;
  const head = ["ITEM", "DETAIL", "AMOUNT"].map((t, i) => ({ text: t, options: { fontFace: F_EN, fontSize: 11, bold: true, color: MUTED, align: i === 2 ? "right" : "left" } }));
  slide.addTable([head], { x, y: top, w, h: 0.45, colW: [3.0, 6.5, 2.5], border: { pt: 0, color: WHITE }, fill: { color: WHITE }, valign: "middle", margin: [4, 8, 4, 8] });
  slide.addShape("line", { x, y: top + 0.45, w, h: 0, line: { color: INK, width: 2 } });

  const rows = $(slideEl).find(".qt-tbl tbody tr").toArray();
  const rowH = 0.95;
  rows.forEach((row, i) => {
    const y = top + 0.45 + i * rowH;
    slide.addShape("line", { x, y: y + rowH, w, h: 0, line: { color: "EEEEEE", width: 1 } });
    slide.addText(`■  ${$(row).find(".qt-cat").text().trim()}`, { x: x + 0.15, y, w: 2.85, h: rowH, valign: "middle", fontFace: F_JP, fontSize: 15, bold: true, color: INK });
    slide.addText($(row).find(".qt-det").text().trim(), { x: x + 3.1, y, w: 6.3, h: rowH, valign: "middle", fontFace: F_JP, fontSize: 11, color: BODY2, lineSpacingMultiple: 1.3 });
    slide.addText($(row).find(".qt-amt").text().trim(), { x: x + 9.4, y, w: 2.6, h: rowH, valign: "middle", align: "right", fontFace: F_EN, fontSize: 18, bold: true, color: RED });
  });

  const totalY = top + 0.45 + rows.length * rowH + 0.25;
  slide.addShape("roundRect", { x: x + 6.5, y: totalY, w: 5.5, h: 0.75, fill: { color: "FBF4F5" }, line: { type: "none" }, rectRadius: 0.06 });
  slide.addText($(slideEl).find(".qt-total .lbl").text().trim(), { x: x + 6.8, y: totalY, w: 2.0, h: 0.75, valign: "middle", fontFace: F_JP, fontSize: 14, bold: true, color: RED });
  slide.addText($(slideEl).find(".qt-total .v").text().trim(), { x: x + 8.5, y: totalY, w: 3.2, h: 0.75, valign: "middle", align: "right", fontFace: F_EN, fontSize: 24, bold: true, color: RED });
  slide.addText($(slideEl).find(".c-body > p").last().text().trim(), { x, y: totalY + 0.95, w, h: 0.35, fontFace: F_JP, fontSize: 9.5, color: MUTED });
  footer(slide, ctx.pageNo(), ctx.total);
}

// ---------- dispatcher ----------

function classify($, slideEl) {
  const html = $.html(slideEl);
  const cls = $(slideEl).attr("class") || "";
  if (cls.includes("cover")) return "cover";
  if (cls.includes("section")) return "section";
  if (html.includes("yoken")) return "yoken";
  if (html.includes("cmp2-body")) return "cmp2";
  if (html.includes("rl-body")) return "rolelist";
  if (html.includes("spk2-body")) return "speakers";
  if (html.includes("ro2-body")) return "runofshow";
  if (html.includes("tk-body")) return "talkqa";
  if (html.includes("ai-body")) return "attractioncards";
  if (html.includes("ad-body")) return "attractiondetail";
  if (html.includes("pf-body")) return "promoteflow";
  if (html.includes("sc-tbl") || html.includes("sc-body")) return "schedule";
  if (html.includes("qt-tbl")) return "quote";
  return "unknown";
}

const GENERATORS = {
  yoken: genYoken,
  cmp2: genCmp2,
  rolelist: genRoleList,
  speakers: genSpeakers,
  runofshow: genRunOfShow,
  talkqa: genTalkQa,
  attractioncards: genAttractionCards,
  attractiondetail: genAttractionDetail,
  promoteflow: genPromoteFlow,
  schedule: genSchedule,
  quote: genQuote,
};

/**
 * @param {string} html - full source document
 * @param {object} [options]
 * @param {boolean} [options.allowGenericFallback=true] - if an individual
 *   slide's type isn't recognized, render it with the generic
 *   Puppeteer-based engine instead of skipping it (requires Chromium to
 *   be available; set to false in environments without it).
 */
async function generate(html, options = {}) {
  const $ = cheerio.load(html);
  const slideEls = $(".slide").toArray();
  const pptx = new pptxgen();
  pptx.author = "html2pptx-app";
  pptx.defineLayout({ name: "SRC_ASPECT", width: 13.333, height: 7.5 });
  pptx.layout = "SRC_ASPECT";

  const warnings = [];
  let pageCounter = 0;
  const pageableTypes = new Set(Object.keys(GENERATORS));
  const total = slideEls.filter((el) => pageableTypes.has(classify($, el))).length;
  const ctx = { total, pageNo: () => ++pageCounter };

  for (const slideEl of slideEls) {
    const type = classify($, slideEl);
    if (type === "cover") { genCover($, slideEl, pptx); continue; }
    if (type === "section") { genSection($, slideEl, pptx); continue; }
    if (GENERATORS[type]) { GENERATORS[type]($, slideEl, pptx, ctx); continue; }

    // unknown slide sub-type within an otherwise-known deck: fall back
    // to the generic per-element engine for just this one slide, rather
    // than failing the whole conversion.
    warnings.push(`Unrecognized slide sub-type (falling back to generic engine): ${($(slideEl).attr("class") || "").slice(0, 60)}`);
    if (options.allowGenericFallback === false) continue;
    try {
      const wrapperHtml = `<!DOCTYPE html><html><head>${$("style").toArray().map((s) => $.html(s)).join("")}</head><body>${$.html(slideEl)}</body></html>`;
      const [slideModel] = await extractDom(wrapperHtml, { slideSelectors: [".slide"] });
      if (slideModel) {
        const slide = pptx.addSlide();
        await renderSlideModelIntoPptx(pptx, slide, slideModel, 13.333 / slideModel.widthPx, { warnings });
      }
    } catch (e) {
      warnings.push(`Generic fallback failed for one slide: ${e.message}`);
    }
  }

  return { pptx, warnings };
}

registerTemplate({ name: "materialpr-pr-event-proposal", fingerprint, generate });

module.exports = { fingerprint, generate };
