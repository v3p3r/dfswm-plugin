/**
 * DFSWM Compliance Plugin — Word document extractor.
 *
 * Uses the Word JavaScript API to read the open document into the plain
 * normalised model consumed by the engine (see ../engine/model.js).
 * This file only runs inside the Office task pane (browser).
 */
"use strict";

const { pointsToCm } = require("../engine/model.js");

const ALIGN = {
  "Word.Alignment.left": "left",
  "Word.Alignment.centered": "center",
  "Word.Alignment.right": "right",
  "Word.Alignment.justified": "justified",
};

/**
 * Read the open Word document into a model object.
 * @returns {Promise<object>} raw model (pass to normalizeModel)
 */
async function extractDocumentModel() {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");

    const paragraphs = body.paragraphs;
    paragraphs.load([
      "text", "styleBuiltIn", "alignment", "leftIndent", "firstLineIndent",
      "lineSpacing", "keepWithNext", "keepPrevious",
    ]);
    paragraphs.font.load(["name", "size", "bold", "italic", "underline", "allCaps"]);

    const sections = context.document.sections;
    // Load section page setup + primary header/footer proxies up-front so
    // their .text is available after the first sync (Word JS requires
    // loading a property on the proxy before reading it post-sync).
    const hf = []; // { sectionIndex, header?, footer?, firstPageHeader?, firstPageFooter? }
    for (let s = 0; s < sections.items.length; s++) {
      const section = sections.items[s];
      section.load("pageSetup");
      const pu = section.pageSetup;
      pu.load(["leftMargin", "rightMargin", "topMargin", "bottomMargin", "headerDistance", "footerDistance", "differentFirstPageHeaderFooter"]);
      const entry = { sectionIndex: s, header: null, footer: null };
      // Primary headers/footers always exist, so no guard is needed.
      entry.header = section.getHeader(Word.HeaderFooterType.primary);
      entry.header.load("text");
      entry.footer = section.getFooter(Word.HeaderFooterType.primary);
      entry.footer.load("text");
      hf.push(entry);
    }

    const tables = body.tables;
    tables.load("rowCount");
    for (const t of tables.items) {
      t.load("rowCount");
      for (const r of t.rows.items) {
        r.load("cellCount");
        for (const c of r.cells.items) {
          c.body.load("text");
          c.body.paragraphs.load("text");
          c.body.paragraphs.font.load("size");
        }
      }
    }

    await context.sync();

    // First-page headers/footers only exist when "different first page" is
    // enabled; getHeader(firstPage) otherwise throws at sync time (the
    // throw surfaces when the queued command runs, so try/catch around the
    // call is useless). Gate the requests behind the loaded flag.
    let needsFirstPageSync = false;
    for (let s = 0; s < sections.items.length; s++) {
      const section = sections.items[s];
      if (!section.pageSetup.differentFirstPageHeaderFooter) continue;
      needsFirstPageSync = true;
      const entry = hf[s];
      entry.firstPageHeader = section.getHeader(Word.HeaderFooterType.firstPage);
      entry.firstPageHeader.load("text");
      entry.firstPageFooter = section.getFooter(Word.HeaderFooterType.firstPage);
      entry.firstPageFooter.load("text");
    }
    if (needsFirstPageSync) await context.sync();

    const raw = {
      docType: "general",
      title: null,
      paragraphs: paragraphs.items.map((p) => {
        const f = p.font;
        return {
          text: p.text,
          style: p.styleBuiltIn ? String(p.styleBuiltIn) : null,
          fontName: f.name,
          fontSize: f.size,
          bold: f.bold,
          italic: f.italic,
          underline: !!f.underline,
          allCaps: f.allCaps,
          alignment: ALIGN[p.alignment] || "left",
          firstLineIndentCm: p.firstLineIndent != null ? pointsToCm(p.firstLineIndent) : null,
          blockIndentCm: p.leftIndent != null ? pointsToCm(p.leftIndent) : null,
          lineSpacing: p.lineSpacing === 1.0 ? "single" : p.lineSpacing === 2.0 ? "double" : p.lineSpacing === 1.5 ? "oneAndHalf" : p.lineSpacing,
          keepWithNext: p.keepWithNext,
          keepPrevious: p.keepPrevious,
          headingLevel: null, // resolved by heuristics below
          pageBreakBefore: false,
        };
      }),
      sections: sections.items.map((s, si) => {
        const pu = s.pageSetup;
        const entry = hf[si];
        return {
          margins: {
            left: pu.leftMargin != null ? pointsToCm(pu.leftMargin) : null,
            right: pu.rightMargin != null ? pointsToCm(pu.rightMargin) : null,
            top: pu.topMargin != null ? pointsToCm(pu.topMargin) : null,
            bottom: pu.bottomMargin != null ? pointsToCm(pu.bottomMargin) : null,
          },
          headerDistanceCm: pu.headerDistance != null ? pointsToCm(pu.headerDistance) : null,
          footerDistanceCm: pu.footerDistance != null ? pointsToCm(pu.footerDistance) : null,
          headers: {
            primary: entry && entry.header ? [entry.header.text || ""] : [],
            firstPage: entry && entry.firstPageHeader ? [entry.firstPageHeader.text || ""] : [],
            evenPages: [],
          },
          footers: {
            primary: entry && entry.footer ? [entry.footer.text || ""] : [],
            firstPage: entry && entry.firstPageFooter ? [entry.firstPageFooter.text || ""] : [],
            evenPages: [],
          },
        };
      }),
      tables: tables.items.map((t) => ({
        headerRow: t.rows.items[0] ? t.rows.items[0].cells.items.map((c) => c.body.text) : [],
        rows: t.rows.items.map((r) => r.cells.items.map((c) => {
          const sizes = c.body.paragraphs.items.map((cp) => cp.font.size).filter((n) => n != null);
          return { text: c.body.text, fontSize: sizes.length ? sizes[0] : null };
        })),
      })),
      footnotes: [], // Word JS does not expose footnotes reliably; handled as manual
    };

    // Heading-level heuristic: bold + centered + short + all-caps => subject/main;
    // bold + caps + left => group; bold + title case => paragraph heading.
    raw.paragraphs.forEach((p) => {
      const t = p.text.trim();
      if (!t) return;
      const short = t.length <= 90;
      if (!p.bold || !short) return;
      const allCaps = t === t.toUpperCase();
      if (p.alignment === "center" && allCaps) p.headingLevel = "subject";
      else if (allCaps && p.alignment !== "center") p.headingLevel = "group";
      else if (/^[A-Z][a-z]/.test(t)) p.headingLevel = "paragraph";
    });

    return raw;
  });
}

module.exports = { extractDocumentModel };
