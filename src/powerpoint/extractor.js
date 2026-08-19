/**
 * DFSWM Compliance Plugin — PowerPoint presentation extractor.
 *
 * Maps the open presentation into the same plain, serialisable model the
 * Word extractor produces (see ../engine/model.js), so the engine and its
 * evaluators are host-agnostic. Each text-bearing shape's paragraphs become
 * model paragraphs tagged with their slideIndex.
 *
 * This file only runs inside the Office task pane (browser). The pure
 * mapping function is exported so it can be unit-tested in Node.
 */
"use strict";

const PPT_ALIGN = {
  "Left": "left",
  "Center": "center",
  "Right": "right",
  "Justify": "justified",
  "Distribute": "justified",
  "JustifyLow": "justified",
  "ThaiDistribute": "justified",
};

/**
 * Pure mapping: plain slide data -> raw model (same shape as the Word
 * extractor's output, consumable directly by the engine).
 *
 * slidesData: [{ name, shapes: [{ name, hasTextFrame,
 *   paragraphs: [{ text, fontName, fontSize, bold, italic, underline,
 *                  alignment }] }] }]
 */
function mapPresentationToModel(slidesData) {
  const paragraphs = [];
  (slidesData || []).forEach((slide, slideIndex) => {
    (slide.shapes || []).forEach((shape) => {
      if (!shape.hasTextFrame) return;
      (shape.paragraphs || []).forEach((p) => {
        const t = String(p.text || "").trim();
        if (!t) return;
        paragraphs.push({
          text: t,
          style: null,
          fontName: p.fontName || null,
          fontSize: p.fontSize ?? null,
          bold: !!p.bold,
          italic: !!p.italic,
          underline: !!p.underline,
          allCaps: false, // PPT font.allCaps is not reliably exposed; evaluators compare text
          alignment: PPT_ALIGN[p.alignment] || "left",
          firstLineIndentCm: null,
          blockIndentCm: null,
          lineSpacing: null,
          keepWithNext: false,
          keepPrevious: false,
          headingLevel: null,
          pageBreakBefore: false,
          slideIndex,
        });
      });
    });
  });

  // Heading heuristic (mirrors the Word extractor): bold + short text.
  paragraphs.forEach((p) => {
    const t = p.text.trim();
    if (!t) return;
    if (t.length > 90 || !p.bold) return;
    const allCaps = t === t.toUpperCase();
    if (p.alignment === "center" && allCaps) p.headingLevel = "subject";
    else if (allCaps && p.alignment !== "center") p.headingLevel = "group";
    else if (/^[A-Z][a-z]/.test(t)) p.headingLevel = "paragraph";
  });

  return {
    docType: "presentation",
    title: null,
    paragraphs: paragraphs.map((p, i) => ({ ...p, index: i })),
    sections: [], // PowerPoint has no page setup; layout rules do not apply
    tables: [],
    footnotes: [],
  };
}

/**
 * Read the open presentation into a model object via the PowerPoint JS API.
 * @returns {Promise<object>} raw model (consumed directly by runCompliance)
 */
async function extractPresentationModel() {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    const slideData = [];
    const jobs = []; // { slide, shape, paragraphs }

    for (let s = 0; s < slides.items.length; s++) {
      const slide = slides.items[s];
      const shapes = slide.shapes;
      const shapesData = [];
      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        shape.load("name");
        shape.load("hasTextFrame");
        shapesData.push({ name: null, hasTextFrame: false, paragraphs: [] });
        jobs.push({ slideIndex: s, shape, shapeData: shapesData[shapesData.length - 1] });
      }
      slideData.push({ name: null, shapes: shapesData });
    }
    await context.sync();

    // Second pass: for shapes with a text frame, load paragraph text and
    // per-paragraph font/alignment via path-based loads so every item's
    // child properties are populated after the sync.
    for (const job of jobs) {
      if (!job.shape.hasTextFrame) continue;
      const textRange = job.shape.textFrame.textRange;
      const paras = textRange.paragraphs;
      paras.load("text, font/name, font/size, font/bold, font/italic, font/underline, paragraphFormat/alignment");
      job.paras = paras;
    }
    await context.sync();

    // Read the loaded values into plain data.
    for (const job of jobs) {
      job.shapeData.name = job.shape.name;
      job.shapeData.hasTextFrame = job.shape.hasTextFrame;
      if (!job.paras) continue;
      for (const p of job.paras.items) {
        job.shapeData.paragraphs.push({
          text: p.text,
          fontName: p.font.name,
          fontSize: p.font.size,
          bold: p.font.bold,
          italic: p.font.italic,
          underline: p.font.underline,
          alignment: p.paragraphFormat.alignment,
        });
      }
    }

    return mapPresentationToModel(slideData);
  });
}

module.exports = { extractPresentationModel, mapPresentationToModel };
