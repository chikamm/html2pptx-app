"use strict";

/**
 * Exercises generateGenericPptx() WITHOUT Puppeteer/Chromium, using a
 * hand-written "extracted DOM model" fixture in the exact shape
 * extractDom() produces.
 *
 * Why this exists: sandboxed/offline dev environments frequently can't
 * download a Chromium binary (blocked egress to storage.googleapis.com).
 * This test isolates and validates the generator half of the pipeline
 * independently of the browser-dependent extractor half, so the two can
 * be developed/verified separately. Run the real end-to-end path with
 * `npm run test:sample` once Chromium is available (e.g. inside the
 * Docker image built from the provided Dockerfile).
 */

const fs = require("fs");
const path = require("path");
const { generateGenericPptx } = require("../server/generator/genericGenerator");

const sampleImgPath = path.join(__dirname, "fixtures", "sample.jpg");
const sampleImgDataUri = `data:image/jpeg;base64,${fs.readFileSync(sampleImgPath).toString("base64")}`;

const slidesModel = [
  {
    // cover slide: red gradient bg, big title with one emphasized word,
    // an eyebrow label, a thin rule shape, a footer image.
    widthPx: 1920,
    heightPx: 1080,
    backgroundColor: "rgb(181, 0, 18)",
    backgroundImage: "linear-gradient(135deg, rgb(181, 0, 18) 0%, rgb(255, 13, 20) 100%)",
    elements: [
      {
        type: "text",
        rect: { x: 130, y: 220, w: 1400, h: 60 },
        text: "SAMPLE PROPOSAL",
        runs: [{ text: "SAMPLE PROPOSAL", style: { color: "rgb(255,255,255)" } }],
        style: { color: "rgb(255,255,255)", fontSizePx: 30, fontWeight: "700", fontStyle: "normal", fontFamily: "Arial", textAlign: "left", lineHeight: "normal" },
        zIndex: 0, order: 0,
      },
      {
        type: "text",
        rect: { x: 130, y: 290, w: 1500, h: 140 },
        text: "これはテストです。強調ワードを含みます。",
        runs: [
          { text: "これはテストです。", style: { color: "rgb(255,255,255)", bold: true } },
          { text: "強調ワード", style: { color: "rgb(255,217,219)", bold: true } },
          { text: "を含みます。", style: { color: "rgb(255,255,255)", bold: true } },
        ],
        style: { color: "rgb(255,255,255)", fontSizePx: 54, fontWeight: "900", fontStyle: "normal", fontFamily: "Noto Sans JP", textAlign: "left", lineHeight: "1.2" },
        zIndex: 0, order: 1,
      },
      {
        type: "box",
        shape: "roundRect",
        rect: { x: 130, y: 470, w: 100, h: 8 },
        style: { backgroundColor: "rgb(255,255,255)", borderWidth: 0, borderStyle: "none", borderRadiusPx: 4 },
        zIndex: 0, order: 2,
      },
      {
        type: "image",
        rect: { x: 1250, y: 700, w: 500, h: 333 },
        src: sampleImgDataUri,
        zIndex: 0, order: 3,
      },
    ],
  },
  {
    // content slide: white bg, heading, a card (box) with a pill (box)
    // and body text on top of it - exercises layered box+text and
    // overlap-based confidence scoring.
    widthPx: 1920,
    heightPx: 1080,
    backgroundColor: "rgb(255,255,255)",
    backgroundImage: null,
    elements: [
      {
        type: "text",
        rect: { x: 96, y: 94, w: 1200, h: 70 },
        text: "本文スライドの見出し",
        runs: [{ text: "本文スライドの見出し", style: { color: "rgb(26,26,26)" } }],
        style: { color: "rgb(26,26,26)", fontSizePx: 54, fontWeight: "900", fontStyle: "normal", fontFamily: "Noto Sans JP", textAlign: "left", lineHeight: "1.1" },
        zIndex: 0, order: 0,
      },
      {
        type: "box",
        shape: "roundRect",
        rect: { x: 96, y: 300, w: 1728, h: 200 },
        style: { backgroundColor: "rgb(250,247,247)", borderWidth: 1, borderStyle: "solid", borderColor: "rgb(240,230,230)", borderRadiusPx: 18 },
        zIndex: 0, order: 1,
      },
      {
        type: "box",
        shape: "roundRect",
        rect: { x: 130, y: 330, w: 180, h: 50 },
        style: { backgroundColor: "rgb(181,0,18)", borderWidth: 0, borderStyle: "none", borderRadiusPx: 25 },
        zIndex: 1, order: 2,
      },
      {
        type: "text",
        rect: { x: 130, y: 330, w: 180, h: 50 },
        text: "PART 1",
        runs: [{ text: "PART 1", style: { color: "rgb(255,255,255)", bold: true } }],
        style: { color: "rgb(255,255,255)", fontSizePx: 20, fontWeight: "700", fontStyle: "normal", fontFamily: "Arial", textAlign: "center", lineHeight: "normal" },
        zIndex: 2, order: 3,
      },
      {
        type: "text",
        rect: { x: 130, y: 410, w: 1600, h: 70 },
        text: "カードの本文テキストがここに入ります。",
        runs: [{ text: "カードの本文テキストがここに入ります。", style: { color: "rgb(85,85,85)" } }],
        style: { color: "rgb(85,85,85)", fontSizePx: 20, fontWeight: "400", fontStyle: "normal", fontFamily: "Noto Sans JP", textAlign: "left", lineHeight: "1.5" },
        zIndex: 1, order: 4,
      },
    ],
  },
];

(async () => {
  const { pptx, lowConfidenceSlides, warnings } = await generateGenericPptx(slidesModel, {});
  const outPath = path.join(__dirname, "..", "tmp_mock_output.pptx");
  await pptx.writeFile({ fileName: outPath });
  console.log("Wrote:", outPath);
  console.log("lowConfidenceSlides:", lowConfidenceSlides);
  console.log("warnings:", warnings);
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
