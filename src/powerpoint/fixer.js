/**
 * DFSWM Compliance Plugin — PowerPoint autofix layer.
 *
 * Applies the subset of fix actions that make sense on slides using the
 * PowerPoint JavaScript API. Paragraph-level fixes target exactly the
 * flagged model paragraphs (via fix.locations), mirroring the extractor's
 * slide → shape → paragraph ordering and skipping blank lines. Word-only
 * actions (margins, page numbers, line spacing, keep-with-next, copy
 * numbers) return ok:false with instructions.
 */
"use strict";

const PPT_ALIGN = {
  left: "Left",
  center: "Center",
  right: "Right",
  justified: "Justify",
};

const TITLE_CASE_MINOR = new Set([
  "of", "and", "to", "the", "from", "a", "an", "in", "on", "at", "for", "with", "by",
]);

function titleCase(s) {
  return s.split(/\s+/).map((word, i) => {
    const lower = word.toLowerCase();
    if (i > 0 && TITLE_CASE_MINOR.has(lower)) return lower;
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  }).join(" ");
}

/** Paragraph indices a fix should target, or null when there are none. */
function paragraphTargetSet(fix) {
  const locs = (fix.locations || [])
    .map((l) => l.paragraphIndex)
    .filter((n) => Number.isInteger(n));
  return locs.length ? new Set(locs) : null;
}

/**
 * Iterate every text paragraph in the deck in the same order the extractor
 * maps them into the model (slides, then shapes with a text frame, then
 * paragraphs, skipping blank lines) and call `cb(paragraph, modelIndex)`.
 * One PowerPoint.run batch is used for the whole traversal.
 */
async function forEachTextParagraph(cb) {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    const shapes = [];
    for (const slide of slides.items) {
      for (const shape of slide.shapes.items) {
        shape.load("hasTextFrame");
        shapes.push(shape);
      }
    }
    await context.sync();

    const paraLists = [];
    for (const shape of shapes) {
      if (!shape.hasTextFrame) continue;
      const paras = shape.textFrame.textRange.paragraphs;
      paras.load("text, font/name, font/size, font/bold, font/italic, font/underline, paragraphFormat/alignment");
      paraLists.push(paras);
    }
    await context.sync();

    let index = 0;
    for (const paras of paraLists) {
      for (const p of paras.items) {
        const t = String(p.text || "").trim();
        if (!t) continue;
        await cb(p, index++);
      }
    }
    await context.sync();
  });
}

/**
 * Apply a font-level mutation to the targeted paragraphs.
 * @param {function} applyFont (font, paragraph) => void
 */
async function mutateTargetedText(fix, applyFont) {
  const targets = paragraphTargetSet(fix);
  let count = 0;
  await forEachTextParagraph((p, index) => {
    if (targets && !targets.has(index)) return;
    count += 1;
    applyFont(p.font, p);
  });
  return count;
}

/**
 * Apply a fix object produced by an evaluator.
 * @returns {Promise<{ok: boolean, message: string}>}
 */
async function applyFix(fix) {
  if (!fix || !fix.action) return { ok: false, message: "No fix provided." };
  try {
    switch (fix.action) {
      case "setFont": {
        const n = await mutateTargetedText(fix, (font) => { font.name = fix.params.fontName; });
        return { ok: true, message: `Font set to ${fix.params.fontName} on ${n} paragraph(s).` };
      }
      case "setFontSize": {
        const n = await mutateTargetedText(fix, (font) => { font.size = fix.params.size; });
        return { ok: true, message: `Font size set to ${fix.params.size} pt on ${n} paragraph(s).` };
      }
      case "setBold": {
        const n = await mutateTargetedText(fix, (font) => { font.bold = fix.params.bold; });
        return { ok: true, message: `Bold ${fix.params.bold ? "applied" : "removed"} on ${n} paragraph(s).` };
      }
      case "setItalic": {
        const n = await mutateTargetedText(fix, (font) => { font.italic = fix.params.italic; });
        return { ok: true, message: `Italic ${fix.params.italic ? "applied" : "removed"} on ${n} paragraph(s).` };
      }
      case "removeUnderline": {
        const n = await mutateTargetedText(fix, (font) => { font.underline = "None"; });
        return { ok: true, message: `Underlining removed from ${n} paragraph(s).` };
      }
      case "setAlignment": {
        const targets = paragraphTargetSet(fix);
        let count = 0;
        await forEachTextParagraph((p, index) => {
          if (targets && !targets.has(index)) return;
          count += 1;
          p.paragraphFormat.alignment = PPT_ALIGN[fix.params.alignment] || "Left";
        });
        return { ok: true, message: `Alignment set to ${fix.params.alignment} on ${count} paragraph(s).` };
      }
      case "fixHeading": {
        const params = fix.params;
        const targets = paragraphTargetSet(fix);
        let count = 0;
        await forEachTextParagraph((p, index) => {
          if (targets && !targets.has(index)) return;
          count += 1;
          let text = p.text;
          let changed = false;
          if (params.caps && text !== text.toUpperCase()) { text = text.toUpperCase(); changed = true; }
          else if (params.titleCase) {
            const tc = titleCase(text);
            if (tc !== text) { text = tc; changed = true; }
          }
          if (params.noTrailingFullStop && /\.\s*$/.test(text)) { text = text.replace(/\.\s*$/, ""); changed = true; }
          if (params.trailingFullStop && !/\.\s*$/.test(text)) { text = text.trimEnd() + "."; changed = true; }
          if (changed) p.text = text;
          if (params.alignment) p.paragraphFormat.alignment = PPT_ALIGN[params.alignment] || "Left";
          if (params.bold) p.font.bold = true;
          if (params.size) p.font.size = params.size;
        });
        return { ok: true, message: `Formatted ${count} heading(s).` };
      }
      case "applyRegexFix": {
        const re = new RegExp(fix.params.pattern, "gm");
        const replacement = typeof fix.params.replacement === "string" ? fix.params.replacement : "";
        let fixed = 0;
        await forEachTextParagraph((p) => {
          if (!re.test(p.text)) return;
          re.lastIndex = 0;
          const newText = p.text.replace(re, replacement);
          if (newText !== p.text) { p.text = newText; fixed += 1; }
        });
        return { ok: true, message: `Applied text fix to ${fixed} paragraph(s).` };
      }
      default:
        return { ok: false, message: `Fix '${fix.action}' is not supported in PowerPoint — apply it manually.` };
    }
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

module.exports = { applyFix };