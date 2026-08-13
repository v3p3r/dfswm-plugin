/**
 * DFSWM Compliance Plugin — autofix layer.
 *
 * Applies simple, safe fixes to the open Word document. Each fix is an
 * object produced by an evaluator: { action, params }.
 */
"use strict";

const ALIGN = {
  left: "Word.Alignment.left",
  center: "Word.Alignment.centered",
  right: "Word.Alignment.right",
  justified: "Word.Alignment.justified",
};

async function applyFix(fix) {
  if (!fix || !fix.action) return { ok: false, message: "No fix provided." };
  try {
    switch (fix.action) {
      case "setFont":
        return await Word.run(async (context) => {
          const body = context.document.body;
          body.font.name = fix.params.fontName;
          await context.sync();
          return { ok: true, message: `Font set to ${fix.params.fontName}.` };
        });
      case "setFontSize":
        return await Word.run(async (context) => {
          const body = context.document.body;
          body.font.size = fix.params.size;
          await context.sync();
          return { ok: true, message: `Body font size set to ${fix.params.size} pt.` };
        });
      case "setAlignment":
        return await Word.run(async (context) => {
          const body = context.document.body;
          body.paragraphs.load("alignment");
          await context.sync();
          for (const p of body.paragraphs.items) {
            p.alignment = ALIGN[fix.params.alignment] || "Word.Alignment.justified";
          }
          await context.sync();
          return { ok: true, message: `Alignment set to ${fix.params.alignment}.` };
        });
      case "setBold":
        return await Word.run(async (context) => {
          const body = context.document.body;
          body.font.bold = fix.params.bold;
          await context.sync();
          return { ok: true, message: `Bold ${fix.params.bold ? "applied" : "removed"}.` };
        });
      case "removeUnderline":
        return await Word.run(async (context) => {
          const body = context.document.body;
          body.font.underline = "Word.UnderlineType.none";
          await context.sync();
          return { ok: true, message: "Underlining removed." };
        });
      case "setLineSpacing":
        return await Word.run(async (context) => {
          const body = context.document.body;
          const spacing = fix.params.spacing === "double" ? 2.0 : fix.params.spacing === "oneAndHalf" ? 1.5 : 1.0;
          for (const p of body.paragraphs.items) {
            p.lineSpacing = spacing;
          }
          await context.sync();
          return { ok: true, message: `Line spacing set to ${fix.params.spacing}.` };
        });
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
      case "keepWithNext":
        return await Word.run(async (context) => {
          const body = context.document.body;
          for (const p of body.paragraphs.items) {
            p.keepWithNext = true;
          }
          await context.sync();
          return { ok: true, message: "Keep-with-next applied to prevent hanging headings." };
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
