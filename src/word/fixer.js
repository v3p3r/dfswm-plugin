/**
 * DFSWM Compliance Plugin — autofix layer.
 *
 * Applies simple, safe fixes to the open Word document. Each fix is an
 * object produced by an evaluator: { action, params, locations }.
 *
 * The engine attaches the finding's locations (paragraph indices in the
 * document body) onto every fix, so paragraph-level fixes here mutate only
 * the flagged paragraphs instead of the whole document. Document-level
 * fixes (margins, page numbers, copy numbers) carry an empty location list
 * and are applied document-wide as before.
 */
"use strict";

const ALIGN = {
  left: "Word.Alignment.left",
  center: "Word.Alignment.centered",
  right: "Word.Alignment.right",
  justified: "Word.Alignment.justified",
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

/**
 * Paragraph indices a fix should target, or null when there are none
 * (document-level fix => apply to every paragraph as a fallback).
 */
function paragraphTargetSet(fix) {
  const locs = (fix.locations || [])
    .map((l) => l.paragraphIndex)
    .filter((n) => Number.isInteger(n));
  return locs.length ? new Set(locs) : null;
}

/**
 * Run a Word batch with the body paragraphs loaded and handed to `fn`.
 * `fn(context, paragraphs)` returns a value that is passed through.
 */
async function withBodyParagraphs(fn) {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load(["text", "font"]);
    await context.sync();
    const result = await fn(context, paragraphs.items);
    await context.sync();
    return result;
  });
}

/** Narrow a collection to the target indices (or keep all when null). */
function pickParagraphs(items, targets) {
  return targets ? items.filter((p, i) => targets.has(i)) : Array.from(items);
}

async function applyFix(fix) {
  if (!fix || !fix.action) return { ok: false, message: "No fix provided." };
  try {
    switch (fix.action) {
      case "setFont":
        return await withBodyParagraphs(async (context, items) => {
          const list = pickParagraphs(items, paragraphTargetSet(fix));
          for (const p of list) p.font.name = fix.params.fontName;
          return { ok: true, message: `Font set to ${fix.params.fontName} on ${list.length} paragraph(s).` };
        });
      case "setFontSize":
        return await withBodyParagraphs(async (context, items) => {
          const list = pickParagraphs(items, paragraphTargetSet(fix));
          for (const p of list) p.font.size = fix.params.size;
          return { ok: true, message: `Font size set to ${fix.params.size} pt on ${list.length} paragraph(s).` };
        });
      case "setAlignment":
        return await withBodyParagraphs(async (context, items) => {
          const list = pickParagraphs(items, paragraphTargetSet(fix));
          for (const p of list) p.alignment = ALIGN[fix.params.alignment] || "Word.Alignment.justified";
          return { ok: true, message: `Alignment set to ${fix.params.alignment} on ${list.length} paragraph(s).` };
        });
      case "setBold":
        return await withBodyParagraphs(async (context, items) => {
          const list = pickParagraphs(items, paragraphTargetSet(fix));
          for (const p of list) p.font.bold = fix.params.bold;
          return { ok: true, message: `Bold ${fix.params.bold ? "applied" : "removed"} on ${list.length} paragraph(s).` };
        });
      case "setItalic":
        return await withBodyParagraphs(async (context, items) => {
          const list = pickParagraphs(items, paragraphTargetSet(fix));
          for (const p of list) p.font.italic = fix.params.italic;
          return { ok: true, message: `Italic ${fix.params.italic ? "applied" : "removed"} on ${list.length} paragraph(s).` };
        });
      case "removeUnderline":
        return await withBodyParagraphs(async (context, items) => {
          const list = pickParagraphs(items, paragraphTargetSet(fix));
          for (const p of list) p.font.underline = "Word.UnderlineType.none";
          return { ok: true, message: `Underlining removed from ${list.length} paragraph(s).` };
        });
      case "setLineSpacing":
        return await withBodyParagraphs(async (context, items) => {
          const spacing = fix.params.spacing === "double" ? 2.0 : fix.params.spacing === "oneAndHalf" ? 1.5 : 1.0;
          const list = pickParagraphs(items, paragraphTargetSet(fix));
          for (const p of list) p.lineSpacing = spacing;
          return { ok: true, message: `Line spacing set to ${fix.params.spacing} on ${list.length} paragraph(s).` };
        });
      case "keepWithNext":
        return await withBodyParagraphs(async (context, items) => {
          const list = pickParagraphs(items, paragraphTargetSet(fix));
          for (const p of list) {
            if (fix.params.keepWithNext) p.keepWithNext = true;
            if (fix.params.keepWithPrevious) p.keepPrevious = true;
          }
          return { ok: true, message: `Keep-together applied to ${list.length} paragraph(s).` };
        });
      case "fixHeading": {
        const params = fix.params;
        return await withBodyParagraphs(async (context, items) => {
          const list = pickParagraphs(items, paragraphTargetSet(fix));
          for (const p of list) {
            let text = p.text;
            let changed = false;
            if (params.caps && text !== text.toUpperCase()) { text = text.toUpperCase(); changed = true; }
            else if (params.titleCase) {
              const tc = titleCase(text);
              if (tc !== text) { text = tc; changed = true; }
            }
            if (params.noTrailingFullStop && /\.\s*$/.test(text)) { text = text.replace(/\.\s*$/, ""); changed = true; }
            if (params.trailingFullStop && !/\.\s*$/.test(text)) { text = text.trimEnd() + "."; changed = true; }
            if (changed) p.insertText(text, "Replace");
            if (params.alignment) p.alignment = ALIGN[params.alignment] || "Word.Alignment.justified";
            if (params.bold) p.font.bold = true;
            if (params.size) p.font.size = params.size;
          }
          return { ok: true, message: `Formatted ${list.length} heading(s).` };
        });
      }
      case "applyRegexFix": {
        const re = new RegExp(fix.params.pattern, "gm");
        const replacement = typeof fix.params.replacement === "string" ? fix.params.replacement : "";
        return await withBodyParagraphs(async (context, items) => {
          let fixed = 0;
          for (const p of items) {
            if (!re.test(p.text)) continue;
            re.lastIndex = 0;
            const newText = p.text.replace(re, replacement);
            if (newText !== p.text) { p.insertText(newText, "Replace"); fixed += 1; }
          }
          return { ok: true, message: `Applied text fix to ${fixed} paragraph(s).` };
        });
      }
      case "setMargins":
        return await Word.run(async (context) => {
          const sections = context.document.sections;
          const cmToPt = (cm) => cm * 28.3465;
          for (const s of sections.items) {
            const pu = s.pageSetup;
            if (fix.params.target === "header-footer") {
              pu.headerDistance = cmToPt(fix.params.valueCm);
              pu.footerDistance = cmToPt(fix.params.valueCm);
            } else {
              const sides = Array.isArray(fix.params.sides) ? fix.params.sides : [];
              if (sides.includes("left")) pu.leftMargin = cmToPt(fix.params.valueCm);
              if (sides.includes("right")) pu.rightMargin = cmToPt(fix.params.valueCm);
              if (sides.includes("top")) pu.topMargin = cmToPt(fix.params.valueCm);
              if (sides.includes("bottom")) pu.bottomMargin = cmToPt(fix.params.valueCm);
            }
          }
          await context.sync();
          return {
            ok: true,
            message: fix.params.target === "header-footer"
              ? `Header/footer distance set to ${fix.params.valueCm} cm.`
              : "Margins updated.",
          };
        });
      case "insertPageNumber":
        // Word JS cannot insert a PAGE field directly; instruct the user.
        return { ok: false, message: "Insert the page number manually: Insert tab > Page Number > Bottom Center." };
      case "insertCopyNumber":
        return await Word.run(async (context) => {
          const body = context.document.body;
          const range = body.insertText("\nCopy No 1 of 1 copy\n", "Word.InsertLocation.start");
          range.load("text");
          await context.sync();
          return { ok: true, message: "Copy number placeholder inserted at the top of the document." };
        });
      default:
        return { ok: false, message: `No implementation for fix action '${fix.action}'.` };
    }
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

module.exports = { applyFix };