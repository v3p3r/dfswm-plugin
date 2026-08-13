"use strict";

/**
 * Test helpers for building mock document models.
 * Each paragraph accepts partial fields; sensible defaults are applied.
 */

function para(text, over = {}) {
  return {
    text,
    style: over.style || null,
    fontName: over.fontName ?? "Times New Roman",
    fontSize: over.fontSize ?? 12,
    bold: !!over.bold,
    italic: !!over.italic,
    underline: !!over.underline,
    allCaps: !!over.allCaps,
    alignment: over.alignment || "justified",
    firstLineIndentCm: over.firstLineIndentCm ?? null,
    blockIndentCm: over.blockIndentCm ?? null,
    lineSpacing: over.lineSpacing ?? "single",
    keepWithNext: !!over.keepWithNext,
    keepPrevious: !!over.keepPrevious,
    headingLevel: over.headingLevel || null,
    pageBreakBefore: !!over.pageBreakBefore,
  };
}

function model(paragraphs, over = {}) {
  return {
    docType: over.docType || "general",
    title: over.title || null,
    paragraphs: paragraphs.map((p, i) => ({ ...p, index: i })),
    sections: over.sections || [
      {
        margins: { left: 2, right: 2, top: 2, bottom: 2 },
        headerDistanceCm: 1.5,
        footerDistanceCm: 1.5,
        headers: { primary: [], firstPage: [], evenPages: [] },
        footers: { primary: [], firstPage: [], evenPages: [] },
      },
    ],
    tables: over.tables || [],
    footnotes: over.footnotes || [],
    get bodyText() {
      return this.paragraphs.map((p) => p.text).join("\n");
    },
  };
}

function servicePaperModel(over = {}) {
  const paragraphs = [
    para("MOMBASA ISLAND ROUTE", { bold: true, allCaps: true, alignment: "center", headingLevel: "subject" }),
    para("Reference:", { bold: true }),
    para("A. JCSC/TRG/23 dated 1 Jul 23"),
    para("INTRODUCTION", { bold: true, allCaps: true, headingLevel: "main" }),
    para("The tank Bde was originally designed to support the infantry."),
    para("AIM", { bold: true, allCaps: true, headingLevel: "main" }),
    para("The aim of this service paper is to examine the employment of the tank Bde."),
    para("PRINCIPLES OF EMPLOYMENT", { bold: true, allCaps: true, headingLevel: "main" }),
    para("General.", { bold: true, headingLevel: "paragraph" }),
    para("The Bde will share the task of breaking through the enemy defences."),
    para("CONCLUSION", { bold: true, allCaps: true, headingLevel: "main" }),
    para("The principles of employment of the tank Bde are sound."),
    para("RECOMMENDATIONS", { bold: true, allCaps: true, headingLevel: "main" }),
    para("It is recommended that the Bde adopts mobile defence."),
    para("V OJWANG", { bold: true, allCaps: true }),
    para("Lt Col", { bold: true }),
    para("SO1 Ops & Trg"),
  ];
  return model(paragraphs, { docType: "service-paper", ...over });
}

module.exports = { para, model, servicePaperModel };
