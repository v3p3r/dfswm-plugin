/**
 * DFSWM Compliance Plugin — document model.
 *
 * The Word task pane extracts the open document into this plain, serialisable
 * model via the Word JavaScript API (see ../word/extractor.js). The engine
 * (evaluators.js / engine.js) only ever sees this model, so the same code runs
 * in Node for tests and in the browser for live checking.
 *
 * Paragraph fields:
 *   text            plain text of the paragraph
 *   style           Word style name, if any ("Normal", "Heading 1", ...)
 *   fontName        e.g. "Times New Roman"
 *   fontSize        in points
 *   bold / italic / underline   booleans (underline: any underline)
 *   allCaps         whether the run is formatted ALL CAPS (Word property)
 *   alignment       "left" | "center" | "right" | "justified"
 *   firstLineIndentCm / blockIndentCm
 *   lineSpacing     "single" | "double" | "oneAndHalf" | numeric multiplier
 *   keepWithNext    paragraph.keepWithNext
 *   keepPrevious    paragraph.keepPrevious
 *   headingLevel    null | "subject" | "main" | "group" | "paragraph"
 *   pageBreakBefore true when the paragraph starts a new page
 *
 * Section fields:
 *   margins   { left, right, top, bottom } in cm
 *   headerDistanceCm / footerDistanceCm
 *   headers / footers — { primary, firstPage, evenPages } text arrays
 *
 * Top-level:
 *   docType   selected document type (see dfswm-rulesets/index.json)
 *   bodyText  convenience concatenation of all paragraph text
 */
"use strict";

const CM_PER_POINT = 2.54 / 72; // 1 pt = 1/72 in = 2.54/72 cm

function pointsToCm(pt) {
  return typeof pt === "number" ? pt * CM_PER_POINT : null;
}

/**
 * Convert raw extracted data into a validated, normalised model.
 * `raw` may already be a model (idempotent).
 */
function normalizeModel(raw) {
  if (!raw || typeof raw !== "object") {
    throw new Error("model must be an object");
  }
  const paragraphs = (raw.paragraphs || []).map((p, i) => ({
    index: i,
    text: String(p.text ?? ""),
    style: p.style || null,
    fontName: p.fontName || null,
    fontSize: p.fontSize ?? null,
    bold: !!p.bold,
    italic: !!p.italic,
    underline: !!p.underline,
    allCaps: !!p.allCaps,
    alignment: p.alignment || "left",
    firstLineIndentCm: p.firstLineIndentCm ?? null,
    blockIndentCm: p.blockIndentCm ?? null,
    lineSpacing: p.lineSpacing ?? null,
    keepWithNext: !!p.keepWithNext,
    keepPrevious: !!p.keepPrevious,
    headingLevel: p.headingLevel || null,
    pageBreakBefore: !!p.pageBreakBefore,
  }));
  const sections = (raw.sections || []).map((s) => ({
    margins: s.margins || { left: null, right: null, top: null, bottom: null },
    headerDistanceCm: s.headerDistanceCm ?? null,
    footerDistanceCm: s.footerDistanceCm ?? null,
    headers: s.headers || { primary: [], firstPage: [], evenPages: [] },
    footers: s.footers || { primary: [], firstPage: [], evenPages: [] },
  }));
  const tables = (raw.tables || []).map((t) => ({
    headerRow: (t.headerRow || []).map((c) => String(c.text ?? c ?? "")),
    rows: (t.rows || []).map((r) => r.map((c) => ({
      text: String(c.text ?? ""),
      fontSize: c.fontSize ?? null,
    }))),
  }));
  const footnotes = (raw.footnotes || []).map((f) => ({
    text: String(f.text ?? ""),
    fontSize: f.fontSize ?? null,
  }));

  const model = {
    docType: raw.docType || "general",
    paragraphs,
    sections,
    tables,
    footnotes,
    title: raw.title || null,
  };
  model.bodyText = paragraphs.map((p) => p.text).join("\n");
  return model;
}

/** True when the paragraph text is blank/whitespace only. */
function isBlankParagraph(p) {
  return !p.text || p.text.trim().length === 0;
}

/**
 * Resolve which paragraphs a check scope covers. Falls back to all
 * paragraphs for scopes that cannot be resolved precisely.
 */
function paragraphsForScope(model, scope) {
  const all = model.paragraphs;
  switch (scope) {
    case "body":
    case "all":
      return all;
    case "subject-heading":
      return all.filter((p) => p.headingLevel === "subject");
    case "main-heading":
      return all.filter((p) => p.headingLevel === "main");
    case "group-heading":
      return all.filter((p) => p.headingLevel === "group");
    case "paragraph-heading":
      return all.filter((p) => p.headingLevel === "paragraph");
    case "heading":
      return all.filter((p) => p.headingLevel != null);
    case "signature-block": {
      const nonBlank = all.filter((p) => !isBlankParagraph(p));
      return nonBlank.slice(-3);
    }
    case "first-page": {
      const cut = all.findIndex((p) => p.pageBreakBefore && p.index > 0);
      const end = cut === -1 ? Math.min(10, all.length) : cut;
      return all.slice(0, end);
    }
    default:
      return all;
  }
}

/** Plain text of a scope (for regex/element checks). */
function textForScope(model, scope) {
  const parts = paragraphsForScope(model, scope).map((p) => p.text);
  if (scope === "header" || scope === "footer") {
    const secText = [];
    for (const s of model.sections) {
      const bag = scope === "header" ? s.headers : s.footers;
      for (const kind of ["primary", "firstPage", "evenPages"]) {
        secText.push(...(bag[kind] || []));
      }
    }
    return secText.join("\n");
  }
  if (scope === "table") {
    return model.tables
      .map((t) => t.rows.map((r) => r.map((c) => c.text).join(" | ")).join("\n"))
      .join("\n");
  }
  if (scope === "table-column-head") {
    return model.tables.map((t) => t.headerRow.join(" | ")).join("\n");
  }
  if (scope === "footnote") {
    return model.footnotes.map((f) => f.text).join("\n");
  }
  return parts.join("\n");
}

/** All document text including headers, footers, tables and footnotes. */
function fullText(model) {
  const bits = [model.bodyText];
  for (const s of model.sections) {
    for (const bag of [s.headers, s.footers]) {
      for (const kind of ["primary", "firstPage", "evenPages"]) {
        bits.push((bag[kind] || []).join("\n"));
      }
    }
  }
  bits.push(
    model.tables.map((t) => t.rows.map((r) => r.map((c) => c.text).join(" ")).join("\n")).join("\n"),
  );
  bits.push(model.footnotes.map((f) => f.text).join("\n"));
  return bits.join("\n");
}

module.exports = {
  CM_PER_POINT,
  pointsToCm,
  normalizeModel,
  isBlankParagraph,
  paragraphsForScope,
  textForScope,
  fullText,
};
