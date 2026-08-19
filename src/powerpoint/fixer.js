/**
 * DFSWM Compliance Plugin — PowerPoint autofix layer.
 *
 * Applies the subset of fix actions that make sense on slides using the
 * PowerPoint JavaScript API. Word-only actions (margins, page numbers, line
 * spacing, keep-with-next, copy numbers) return ok:false with instructions.
 */
"use strict";

const PPT_ALIGN = {
  left: "Left",
  center: "Center",
  right: "Right",
  justified: "Justify",
};

/**
 * Iterate every text-bearing shape in the deck and apply a font-level
 * mutation queued inside a single PowerPoint.run batch.
 * @param {function} applyFont (font, context) => void
 */
async function mutateAllText(applyFont) {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    const jobs = [];
    for (let s = 0; s < slides.items.length; s++) {
      const shapes = slides.items[s].shapes;
      for (let i = 0; i < shapes.items.length; i++) {
        const shape = shapes.items[i];
        shape.load("hasTextFrame");
        jobs.push(shape);
      }
    }
    await context.sync();
    for (const shape of jobs) {
      if (!shape.hasTextFrame) continue;
      const font = shape.textFrame.textRange.font;
      applyFont(font);
    }
    await context.sync();
  });
}

/**
 * Apply a fix object produced by an evaluator.
 * @returns {Promise<{ok: boolean, message: string}>}
 */
async function applyFix(fix) {
  if (!fix || !fix.action) return { ok: false, message: "No fix provided." };
  try {
    switch (fix.action) {
      case "setFont":
        await mutateAllText((font) => {
          font.name = fix.params.fontName;
        });
        return { ok: true, message: `Font set to ${fix.params.fontName} on all slide text.` };
      case "setFontSize":
        await mutateAllText((font) => {
          font.size = fix.params.size;
        });
        return { ok: true, message: `Font size set to ${fix.params.size} pt on all slide text.` };
      case "setBold":
        await mutateAllText((font) => {
          font.bold = fix.params.bold;
        });
        return { ok: true, message: `Bold ${fix.params.bold ? "applied" : "removed"} on all slide text.` };
      case "removeUnderline":
        await mutateAllText((font) => {
          font.underline = "None";
        });
        return { ok: true, message: "Underlining removed from all slide text." };
      case "setAlignment":
        return await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          const jobs = [];
          for (let s = 0; s < slides.items.length; s++) {
            const shapes = slides.items[s].shapes;
            for (let i = 0; i < shapes.items.length; i++) {
              const shape = shapes.items[i];
              shape.load("hasTextFrame");
              jobs.push(shape);
            }
          }
          await context.sync();
          for (const shape of jobs) {
            if (!shape.hasTextFrame) continue;
            const paras = shape.textFrame.textRange.paragraphs;
            paras.load("text");
            for (const p of paras.items) {
              p.paragraphFormat.alignment = PPT_ALIGN[fix.params.alignment] || "Left";
            }
          }
          await context.sync();
          return { ok: true, message: `Alignment set to ${fix.params.alignment}.` };
        });
      default:
        return { ok: false, message: `Fix '${fix.action}' is not supported in PowerPoint — apply it manually.` };
    }
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

module.exports = { applyFix };
