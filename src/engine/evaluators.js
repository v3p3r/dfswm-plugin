/**
 * DFSWM Compliance Plugin — evaluators.
 *
 * Each evaluator has the signature:
 *   (model, rule, ctx) => Finding[]
 *
 * model  — normalised document model (see model.js)
 * rule   — the ruleset rule object (id, severity, check, autofix, ...)
 * ctx    — { references } optional reference data (e.g. Ch 16 abbreviations)
 *
 * A Finding is:
 *   { ruleId, severity, title, description, chapter, paragraphs,
 *     message, locations: [{ paragraphIndex?, snippet? }],
 *     fix: { action, params } | null, checkable: bool, manual?: bool }
 *
 * An evaluator that cannot be reliably automated returns a single
 * finding with checkable:false (rendered as a checklist item).
 */
"use strict";

const {
  pointsToCm,
  isBlankParagraph,
  paragraphsForScope,
  textForScope,
  fullText,
} = require("./model.js");

const RANKS = /\b(Lt Col|Lieutenant Colonel|Maj|Major|Capt|Captain|Lt|Lieutenant|WO|Warrant Officer|Sgt|Sergeant|Cpl|Corporal|Brig|Brigadier|Col|Colonel|Gen|General|SO1|SO2|SO3)\b/;

function finding(rule, message, opts = {}) {
  return {
    ruleId: rule.id,
    severity: rule.severity || "warning",
    title: rule.title || rule.id,
    description: rule.description || "",
    chapter: rule.chapter || "",
    paragraphs: rule.paragraphs || "",
    message,
    locations: opts.locations || [],
    fix: opts.fix !== undefined ? opts.fix : null,
    checkable: opts.checkable !== false,
  };
}

function uncheckable(rule, message) {
  return { ...finding(rule, message), checkable: false };
}

function escapeRegExp(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

const EVALUATORS = {};

/* ------------------------------------------------------------------ */
/* Core formatting                                                     */
/* ------------------------------------------------------------------ */

EVALUATORS["font-name"] = (model, rule) => {
  const allowed = (rule.check.params.allowed || []).map((f) => f.toLowerCase());
  const scope = paragraphsForScope(model, rule.check.scope);
  const bad = scope.filter((p) => p.text.trim() && p.fontName && !allowed.includes(p.fontName.toLowerCase()));
  if (!bad.length) return [];
  return [finding(rule, `Font '${bad[0].fontName}' is not one of the allowed fonts (${rule.check.params.allowed.join(", ")}).`, {
    locations: bad.map((p) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 80) })),
    fix: { action: "setFont", params: { fontName: rule.check.params.allowed[0] } },
  })];
};

EVALUATORS["font-size"] = (model, rule) => {
  const { size, comparison = "eq" } = rule.check.params;
  const scope = rule.check.scope;

  if (scope === "table") {
    if (!model.tables.length) return [uncheckable(rule, "No tables in the document.")];
    const bad = [];
    for (const t of model.tables) {
      for (const row of t.rows) {
        for (const cell of row) {
          if (cell.fontSize != null && cell.fontSize !== size) bad.push(cell);
        }
      }
    }
    if (!bad.length) return [];
    return [finding(rule, `${bad.length} table cell(s) are not ${size} pt (e.g. ${bad[0].fontSize} pt).`)];
  }

  if (scope === "footnote") {
    if (!model.footnotes.length) return [uncheckable(rule, "No footnote data extracted from the document.")];
    const bad = model.footnotes.filter((f) => f.fontSize != null && f.fontSize !== size);
    if (!bad.length) return [];
    return [finding(rule, `${bad.length} footnote(s) are not ${size} pt.`)];
  }

  const paragraphs = paragraphsForScope(model, scope);
  const bad = paragraphs.filter((p) => p.text.trim() && p.fontSize != null && p.fontSize !== size);
  if (!bad.length) return [];
  return [finding(rule, `${bad.length} item(s) in scope '${scope}' are not ${size} pt (e.g. ${bad[0].fontSize} pt).`, {
    locations: bad.slice(0, 10).map((p) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 80) })),
    fix: { action: "setFontSize", params: { size } },
  })];
};

EVALUATORS["justification"] = (model, rule) => {
  const want = rule.check.params.alignment || "justified";
  const scope = paragraphsForScope(model, rule.check.scope);
  const bad = scope.filter((p) => p.text.trim() && p.alignment && p.alignment !== want);
  if (!bad.length) return [];
  return [finding(rule, `${bad.length} paragraph(s) are not ${want} (e.g. '${bad[0].alignment}').`, {
    locations: bad.slice(0, 10).map((p) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 80) })),
    fix: { action: "setAlignment", params: { alignment: want } },
  })];
};

EVALUATORS["bold"] = (model, rule) => {
  const expected = !!rule.check.params.expected;
  const scope = rule.check.scope;
  if (scope === "table-column-head") {
    return [uncheckable(rule, "Table header-row detection is not implemented; verify column headings are bold manually.")];
  }
  const paragraphs = paragraphsForScope(model, scope);
  const bad = paragraphs.filter((p) => p.text.trim() && p.bold !== expected);
  if (!bad.length) return [];
  return [finding(rule, `${bad.length} item(s) in scope '${scope}' should be ${expected ? "bold" : "not bold"}.`, {
    locations: bad.slice(0, 10).map((p) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 80) })),
    fix: { action: "setBold", params: { bold: expected } },
  })];
};

EVALUATORS["italic"] = (model, rule) => {
  const expected = !!rule.check.params.expected;
  const scope = paragraphsForScope(model, rule.check.scope);
  const bad = scope.filter((p) => p.text.trim() && /[“"']/.test(p.text) && p.italic !== expected);
  if (!bad.length) return [];
  return [finding(rule, "Quoted text or publication titles should be italicised.", {
    locations: bad.slice(0, 10).map((p) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 80) })),
  })];
};

EVALUATORS["underline"] = (model, rule) => {
  const scope = paragraphsForScope(model, rule.check.scope);
  const bad = scope.filter((p) => p.underline);
  if (!bad.length) return [];
  return [finding(rule, `Underlining must not be used for emphasis in typed text (${bad.length} paragraph(s)).`, {
    locations: bad.slice(0, 10).map((p) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 80) })),
    fix: { action: "removeUnderline", params: {} },
  })];
};

EVALUATORS["margin"] = (model, rule) => {
  const params = rule.check.params;
  const section = model.sections[0];
  if (!section) return [uncheckable(rule, "No section layout data available; check page margins manually.")];

  // Header/footer distance variant
  if (params.target === "header-footer") {
    const issues = [];
    if (section.headerDistanceCm != null && Math.abs(section.headerDistanceCm - params.valueCm) > 0.05) {
      issues.push(`header distance is ${section.headerDistanceCm.toFixed(2)} cm (should be ${params.valueCm} cm)`);
    }
    if (section.footerDistanceCm != null && Math.abs(section.footerDistanceCm - params.valueCm) > 0.05) {
      issues.push(`footer distance is ${section.footerDistanceCm.toFixed(2)} cm (should be ${params.valueCm} cm)`);
    }
    if (!issues.length) return [];
    return [finding(rule, issues.join("; ") + ".", {
      fix: { action: "setMargins", params: { sides: [], target: "header-footer", valueCm: params.valueCm } },
    })];
  }

  const m = section.margins || {};
  const issues = [];
  for (const side of params.sides || ["left", "right", "top", "bottom"]) {
    const v = m[side];
    if (v == null) continue;
    if (Math.abs(v - params.valueCm) > 0.05) {
      issues.push(`${side} margin is ${v.toFixed(2)} cm (should be ${params.valueCm} cm)`);
    }
  }
  if (!issues.length) return [];
  return [finding(rule, issues.join("; ") + ".", {
    fix: { action: "setMargins", params: { sides: params.sides || ["left", "right", "top", "bottom"], valueCm: params.valueCm } },
  })];
};

EVALUATORS["indent"] = (model, rule) => {
  const { firstLineCm, blockIndentCm } = rule.check.params;
  const scope = paragraphsForScope(model, rule.check.scope);
  const bad = scope.filter((p) => {
    if (!p.text.trim()) return false;
    if (firstLineCm != null && p.firstLineIndentCm != null && Math.abs(p.firstLineIndentCm - firstLineCm) > 0.05) return true;
    if (blockIndentCm != null && p.blockIndentCm != null && Math.abs(p.blockIndentCm - blockIndentCm) > 0.05) return true;
    return false;
  });
  if (!bad.length) return [];
  return [finding(rule, `${bad.length} paragraph(s) have incorrect indentation.`, {
    locations: bad.slice(0, 10).map((p) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 80) })),
  })];
};

EVALUATORS["line-spacing"] = (model, rule) => {
  const allowed = Array.isArray(rule.check.params.spacing) ? rule.check.params.spacing : [rule.check.params.spacing];
  const scope = paragraphsForScope(model, rule.check.scope);
  const bad = scope.filter((p) => p.text.trim() && p.lineSpacing && !allowed.includes(p.lineSpacing));
  if (!bad.length) return [];
  return [finding(rule, `${bad.length} paragraph(s) use line spacing '${bad[0].lineSpacing}' (expected ${allowed.join(" or ")}).`, {
    locations: bad.slice(0, 10).map((p) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 80) })),
    fix: { action: "setLineSpacing", params: { spacing: allowed[0] } },
  })];
};

EVALUATORS["blank-line"] = (model, rule) => {
  const scope = paragraphsForScope(model, rule.check.scope);
  if (!scope.length) return [];
  const lastNonBlank = [...scope].reverse().find((p) => !isBlankParagraph(p));
  if (!lastNonBlank) return [];
  const after = scope.slice(lastNonBlank.index + 1);
  const hasBlank = after.some((p) => isBlankParagraph(p));
  if (hasBlank) return [];
  return [uncheckable(rule, "No blank line detected between the last line of text and the page number; verify in Print Layout.")];
};

EVALUATORS["tab-stop"] = (model, rule) => {
  return [uncheckable(rule, "Tab stop intervals cannot be read from the document model; verify default tabs are 1.0 cm in Page Setup.")];
};

EVALUATORS["keep-with-next"] = (model, rule) => {
  const { keepWithNext = false, keepWithPrevious = false } = rule.check.params;
  const scope = paragraphsForScope(model, rule.check.scope);
  const bad = scope.filter((p) => p.text.trim() && (keepWithNext ? !p.keepWithNext : !p.keepPrevious));
  if (!bad.length) return [];
  return [finding(rule, `${bad.length} item(s) are at risk of being stranded at a page boundary (hanging heading / orphaned signature).`, {
    locations: bad.slice(0, 10).map((p) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 80) })),
    fix: { action: "keepWithNext", params: { keepWithNext } },
  })];
};

EVALUATORS["page-number-position"] = (model, rule) => {
  const footerText = model.sections.map((s) => (s.footers.primary || []).join(" ")).join(" ").trim();
  if (footerText) return [];
  return [finding(rule, "No page number found in the footer (page numbers go at bottom centre, above the classification).", {
    fix: { action: "insertPageNumber", params: { position: rule.check.params.position || "center" } },
  })];
};

EVALUATORS["page-number-format"] = (model, rule) => {
  const fmt = rule.check.params.format;
  const footerText = model.sections.map((s) => (s.footers.primary || []).join(" ")).join(" ").trim();
  if (fmt === "X of Y") {
    if (/\d+\s*of\s*\d+/.test(footerText)) return [];
    return [finding(rule, "Footer page number must use the 'X of Y' total-pages format for this classification.")];
  }
  if (fmt === "annex-letter-page") {
    if (/^[A-Z]-\d+/.test(footerText)) return [];
    return [finding(rule, "Annex pages must be numbered with the annex letter (e.g. A-1, A-2).")];
  }
  return [uncheckable(rule, `Page numbering format '${fmt}' needs manual verification.`)];
};

EVALUATORS["heading-style"] = (model, rule) => {
  const params = rule.check.params;
  const scope = paragraphsForScope(model, rule.check.scope);
  if (!scope.length) return [];

  // Reference body size for the "+2pt" main-heading rule.
  let bodySize = null;
  if (params.sizeOffsetPoints) {
    const counts = {};
    for (const p of model.paragraphs) {
      if (p.headingLevel) continue;
      if (p.fontSize == null) continue;
      counts[p.fontSize] = (counts[p.fontSize] || 0) + 1;
    }
    const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    bodySize = sorted.length ? Number(sorted[0][0]) : null;
  }

  const issues = [];
  for (const p of scope) {
    const list = [];
    if (params.alignment && p.alignment !== params.alignment) list.push(`alignment '${p.alignment}' (expected ${params.alignment})`);
    if (params.bold && !p.bold) list.push("not bold");
    if (params.caps && !p.allCaps && p.text !== p.text.toUpperCase()) list.push("not all capitals");
    if (params.titleCase && !/^([A-Z][a-z]+[\s,&\-–—]*)+/.test(p.text.trim())) list.push("not title case");
    if (params.noTrailingFullStop && /\.\s*$/.test(p.text.trim())) list.push("has trailing full stop");
    if (params.trailingFullStop && !/\.\s*$/.test(p.text.trim())) list.push("missing trailing full stop");
    if (params.sizeOffsetPoints && bodySize != null && p.fontSize != null && p.fontSize !== bodySize + params.sizeOffsetPoints) {
      list.push(`size ${p.fontSize} pt (expected ${bodySize + params.sizeOffsetPoints} pt)`);
    }
    if (list.length) issues.push({ p, list });
  }
  if (!issues.length) return [];
  return [finding(rule, `${issues.length} heading(s) do not conform: ${issues[0].list.join("; ")}.`, {
    locations: issues.map(({ p }) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 80) })),
    fix: { action: "fixHeading", params },
  })];
};

/* ------------------------------------------------------------------ */
/* Structure                                                           */
/* ------------------------------------------------------------------ */

EVALUATORS["required-element"] = (model, rule) => {
  const params = rule.check.params;
  const text = textForScope(model, rule.check.scope);

  // Keyword list: [{ id, pattern }] — each is a searchable regex.
  if (params.keywords) {
    const missing = params.keywords.filter((k) => !new RegExp(k.pattern || escapeRegExp(k.id), "i").test(text));
    if (!missing.length) return [];
    return [finding(rule, `Missing required element(s): ${missing.map((k) => k.id).join(", ")}.`)];
  }

  if (params.elements) {
    // Abstract concept ids (e.g. "file-reference") cannot be matched as
    // literal text; only verify element ids that are literal phrases
    // (contain a space), and mark abstract structural elements for review.
    const literalPhrases = params.elements.filter((el) => /\s/.test(el));
    if (literalPhrases.length) {
      const missing = literalPhrases.filter((el) => !text.toLowerCase().includes(el.toLowerCase()));
      if (missing.length) return [finding(rule, `Missing required element(s): ${missing.join(", ")}.`)];
      return [];
    }
    return [uncheckable(rule, `Verify ${params.elements.join(", ")} manually (structural elements cannot be auto-detected).`)];
  }

  if (params.keyword) {
    const found = new RegExp(params.keyword, "i").test(text);
    if (found) return [];
    return [finding(rule, `Required element '${params.keyword}' not found in scope '${rule.check.scope}'.`)];
  }

  if (params.rule) {
    return [uncheckable(rule, `Structural requirement '${params.rule}' needs manual verification.`)];
  }
  return [uncheckable(rule, "Required-element check needs a keyword, keywords, elements list, or rule.")];
};

EVALUATORS["element-order"] = (model, rule) => {
  const params = rule.check.params;
  const text = textForScope(model, rule.check.scope);

  if (params.mandatory) {
    const lines = text.split("\n").map((l) => l.trim().replace(/^\d+[.)]\s*/, "").toLowerCase());
    const wanted = params.mandatory.map((m) => m.toLowerCase());
    const positions = wanted.map((w) => lines.findIndex((l) => l === w));
    const missing = wanted.filter((w, i) => positions[i] === -1);
    if (missing.length) {
      return [finding(rule, `Mandatory heading(s) missing: ${missing.join(", ").toUpperCase()}.`)];
    }
    if (params.order !== false) {
      for (let i = 1; i < positions.length; i++) {
        if (positions[i] < positions[i - 1]) {
          return [finding(rule, `Headings out of order: '${params.mandatory[i].toUpperCase()}' appears before '${params.mandatory[i - 1].toUpperCase()}'.`)];
        }
      }
    }
    return [];
  }

  if (params.rule === "intro-body-conclusion") {
    return [uncheckable(rule, "Verify the letter text moves through introduction, body and conclusion.")];
  }
  if (params.rule === "conops-elements" || params.rule === "series-items" || params.rule === "aim-after-introduction") {
    return [uncheckable(rule, `Sequence rule '${params.rule}' needs manual verification.`)];
  }
  return [uncheckable(rule, "Order check needs mandatory headings or a rule.")];
};

// Numbering markers per level (1..5) of the DFSWM scheme.
const NUMBER_MARKERS = [
  { re: /^\d+\.\t/, label: "Arabic numeral + full stop + tab" }, // 1.
  { re: /^[a-z]\.\t/, label: "lower-case letter + full stop + tab" }, // a.
  { re: /^\(\d+\)\t/, label: "Arabic numeral in brackets + tab" }, // (1)
  { re: /^\([ivx]+\)\t/, label: "lower-case Roman in brackets + tab" }, // (a)
  { re: /^[ivx]+\.\t/, label: "lower-case Roman + full stop + tab" }, // i.
];

function levelFromIndent(p) {
  if (p.blockIndentCm == null) return 1;
  return Math.min(5, Math.max(1, Math.round(p.blockIndentCm / 1.0) + 1));
}

EVALUATORS["paragraph-numbering"] = (model, rule) => {
  const params = rule.check.params;
  const scope = paragraphsForScope(model, rule.check.scope);

  if (params.rule === "single-paragraph-not-numbered") {
    const numbered = scope.filter((p) => /^\d+\.\t|^\d+\.\s/.test(p.text.trim()));
    if (scope.length === 1 && numbered.length === 1) {
      return [finding(rule, "A single-paragraph document must not number its paragraph.")];
    }
    return [];
  }
  if (params.rule === "never-numbered" || params.rule === "prefer-un-numbered") {
    const numbered = scope.filter((p) => /^\d+\.\t|^\d+\.\s/.test(p.text.trim()));
    if (numbered.length) {
      return [finding(rule, `${numbered.length} paragraph(s) are numbered; paragraphs in this letter type should not be numbered.`, {
        locations: numbered.map((p) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 60) })),
      })];
    }
    return [];
  }
  if (params.rule === "optional-numbering") return [];
  if (params.rule === "numbered-items-with-headings") {
    const items = scope.filter((p) => /^\d+\./.test(p.text.trim()));
    if (!items.length) return [finding(rule, "Agenda items should be numbered, each with a heading.")];
    return [];
  }

  // Hierarchy check: every numbered paragraph's marker must match its
  // indent level in the 1. a. (1) (a) i. scheme.
  const bad = [];
  for (const p of scope) {
    const t = p.text.trim();
    if (!/^(\d+\.|[a-z]\.|\(\d+\)|\([ivx]+\)|[ivx]+\.)\t/.test(t)) continue;
    const level = levelFromIndent(p);
    const marker = NUMBER_MARKERS[level - 1];
    if (marker && !marker.re.test(t)) {
      bad.push({ p, level, marker });
    }
  }
  if (!bad.length) return [];
  return [finding(rule, `${bad.length} numbered paragraph(s) do not match the numbering scheme (e.g. level ${bad[0].level} requires '${bad[0].marker.label}').`, {
    locations: bad.slice(0, 10).map(({ p }) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 60) })),
  })];
};

EVALUATORS["no-single-subparagraph"] = (model, rule) => {
  const scope = paragraphsForScope(model, rule.check.scope);
  for (let i = 0; i < scope.length; i++) {
    const p = scope[i];
    if (!/^\d+\.\t|^\d+\.\s/.test(p.text.trim())) continue;
    const sub = [];
    for (let j = i + 1; j < scope.length; j++) {
      const q = scope[j];
      if (/^\d+\.\t|^\d+\.\s/.test(q.text.trim())) break;
      if (/^[a-z]\.\t/.test(q.text.trim())) sub.push(q);
    }
    if (sub.length === 1) {
      return [finding(rule, `Paragraph ${p.text.trim().split("\t")[0] || p.text.trim().split(" ")[0]} has a single sub-paragraph; a paragraph must never have exactly one sub-paragraph.`, {
        locations: [{ paragraphIndex: p.index, snippet: p.text.slice(0, 60) }, { paragraphIndex: sub[0].index, snippet: sub[0].text.slice(0, 60) }],
      })];
    }
  }
  return [];
};

EVALUATORS["all-caps"] = (model, rule) => {
  const expected = !!rule.check.params.expected;
  const scope = paragraphsForScope(model, rule.check.scope);
  const bad = scope.filter((p) => p.text.trim() && (expected ? p.text !== p.text.toUpperCase() : p.text === p.text.toUpperCase()));
  if (!bad.length) return [];
  return [finding(rule, `${bad.length} item(s) are not in capital letters as required for signal messages.`, {
    locations: bad.slice(0, 10).map((p) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 80) })),
  })];
};

/* ------------------------------------------------------------------ */
/* Text patterns                                                       */
/* ------------------------------------------------------------------ */

EVALUATORS["regex"] = (model, rule) => {
  const params = rule.check.params;
  if (!params.pattern) return [uncheckable(rule, "No pattern supplied.")];
  const re = new RegExp(params.pattern, "m");
  const text = textForScope(model, rule.check.scope);
  const matches = text.match(re);
  const found = !!matches;
  const expect = params.expect || "present";

  if (expect === "absent" && found) {
    return [finding(rule, `Found forbidden text matching '${params.pattern}': "${String(matches[0]).slice(0, 60)}".`, {
      fix: params.mustBeBold ? null : rule.autofix && rule.autofix.available ? { action: "applyRegexFix", params: { pattern: params.pattern } } : null,
    })];
  }
  if (expect === "present" && !found) {
    return [finding(rule, `Required text matching '${params.pattern}' not found.`)];
  }
  if (found && params.mustBeBold) {
    const markings = /TOP SECRET|SECRET|CONFIDENTIAL|RESTRICTED|IN CONFIDENCE|Flag:|Flags:/;
    const scope = paragraphsForScope(model, rule.check.scope);
    const bad = scope.filter((p) => markings.test(p.text) && !p.bold);
    if (bad.length) {
      return [finding(rule, "Security marking / flag reference found but not in bold.", {
        locations: bad.map((p) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 60) })),
        fix: { action: "setBold", params: { bold: true } },
      })];
    }
  }
  return [];
};

EVALUATORS["date-format"] = (model, rule) => {
  const params = rule.check.params;
  const re = new RegExp(params.pattern, "m");
  const text = textForScope(model, rule.check.scope);
  const found = re.test(text);
  const expect = params.expect || "present";
  if (expect === "absent" && found) return [finding(rule, "Date/time written in a non-conforming format.")];
  if (expect === "present" && !found) return [finding(rule, "No conforming date found in the expected format.")];
  return [];
};

EVALUATORS["time-format"] = EVALUATORS["date-format"];

EVALUATORS["dtg-format"] = (model, rule) => {
  const params = rule.check.params;
  if (params.rule === "dtg-after-signature") {
    return [uncheckable(rule, "Verify the DTG was inserted after the releasing officer signed the message.")];
  }
  if (params.rule === "gmt-for-inter-zone") {
    return [uncheckable(rule, "Verify GMT suffix is used for inter-zone messages or messages to addressees outside Kenya.")];
  }
  const re = new RegExp(params.pattern, "m");
  const found = re.test(fullText(model));
  if (!found) return [finding(rule, "No valid Date-Time Group (DDHHMM zone MON YY) found in the message.")];
  return [];
};

/* ------------------------------------------------------------------ */
/* Security / markings                                                 */
/* ------------------------------------------------------------------ */

EVALUATORS["security-marking"] = (model, rule) => {
  const params = rule.check.params;
  const text = fullText(model);
  const allowed = params.allowed || ["TOP SECRET", "SECRET", "CONFIDENTIAL", "RESTRICTED"];
  const found = allowed.filter((c) => text.includes(c));
  if (params.mustMatchDocumentClassification && found.length === 0) {
    return [finding(rule, "Document classification not found in headers/footers.")];
  }
  if (params.privacyMarking && !text.includes(params.privacyMarking)) {
    return [finding(rule, `Privacy marking '${params.privacyMarking}' must appear at top and bottom of every page.`)];
  }
  if (!found.length) {
    return [finding(rule, "No recognised security classification found on the document.")];
  }
  return [];
};

EVALUATORS["copy-number"] = (model, rule) => {
  const params = rule.check.params;
  const text = fullText(model);
  const pattern = params.pattern || /Copy No \d+ of \d+ copies?/;
  const needs = text.includes("TOP SECRET") || text.includes("SECRET") || params.requiredIf === "order";
  if (needs && !pattern.test(text)) {
    return [finding(rule, "Document must bear a copy number (e.g. 'Copy No 1 of 20 copies') in the top right of the first page.", {
      fix: { action: "insertCopyNumber", params: {} },
    })];
  }
  return [];
};

EVALUATORS["precedence-marking"] = (model, rule) => {
  const params = rule.check.params;
  const allowed = params.allowed || ["IMMEDIATE", "PRIORITY"];
  const full = fullText(model);
  const firstPage = textForScope(model, "first-page");
  const foundAny = allowed.filter((w) => new RegExp(`\\b${w}\\b`).test(full));
  if (!foundAny.length) return []; // precedence is optional
  const inFirstPage = allowed.filter((w) => new RegExp(`\\b${w}\\b`).test(firstPage));
  const misplaced = foundAny.filter((w) => !inFirstPage.includes(w));
  if (misplaced.length) {
    return [finding(rule, `Precedence '${misplaced.join(", ")}' must appear in capitals at the top right of the FIRST page only.`)];
  }
  return [];
};

EVALUATORS["signature-block"] = (model, rule) => {
  const params = rule.check.params || {};
  const scope = paragraphsForScope(model, "signature-block");
  const blockText = scope.map((p) => p.text).join("\n");
  if (params.elements) {
    const missing = [];
    if (params.elements.includes("name") && !/[A-Z][A-Z\s]+$/.test(blockText)) missing.push("name (in block capitals)");
    if (params.elements.includes("rank") && !RANKS.test(blockText)) missing.push("rank");
    if (params.elements.includes("appointment") && !/\b(SO\d|Comd|Cmd|Officer|Adm|Ops|Trg|Staff|Bde|Div)\b/i.test(blockText)) missing.push("appointment");
    if (missing.length) return [finding(rule, `Signature block missing: ${missing.join(", ")}.`)];
    return [];
  }
  if (params.rule) return [uncheckable(rule, "Signature rule needs manual verification.")];
  return [];
};

EVALUATORS["distribution-order"] = (model, rule) => {
  return [uncheckable(rule, "Verify addressees are listed: external action, external information, then internal.")];
};

EVALUATORS["draft-marking"] = (model, rule) => {
  const text = textForScope(model, "first-page");
  if (!new RegExp(rule.check.params.keyword || "DRAFT").test(text)) {
    return [finding(rule, "The word 'DRAFT' must appear at the top of the first page, below the security grading.")];
  }
  return [];
};

EVALUATORS["not-first-person"] = (model, rule) => {
  const params = rule.check.params;
  const scope = paragraphsForScope(model, rule.check.scope);
  const text = scope.map((p) => p.text).join(" ");
  const matches = text.match(/\b(I|me|my|mine|we|our|ours)\b/g) || [];
  if (params.firstPersonAllowed === false && matches.length > 2) {
    return [finding(rule, `First-person pronouns used ${matches.length} times; Service correspondence is normally written in the third person.`, {
      locations: scope.filter((p) => /\b(I|me|my|we|our)\b/.test(p.text)).slice(0, 5).map((p) => ({ paragraphIndex: p.index, snippet: p.text.slice(0, 80) })),
    })];
  }
  if (params.firstPersonAllowed === true && !matches.length) {
    return [finding(rule, "DO letters should be written in the first person.")];
  }
  return [];
};

EVALUATORS["abbreviation-usage"] = (model, rule) => {
  const params = rule.check.params;
  const text = textForScope(model, rule.check.scope);
  if (params.rule === "explained-on-first-use") {
    const acro = text.match(/\b[A-Z]{3,}\b/g) || [];
    const unexplained = acro.filter((a) => !new RegExp(`[A-Za-z]+\\s*\\(${a}\\)|\\(${a}\\)`).test(text));
    if (unexplained.length) {
      return [finding(rule, `Abbreviation(s) not explained on first use: ${[...new Set(unexplained)].slice(0, 6).join(", ")}.`)];
    }
    return [];
  }
  if (params.rule === "consistent-use") {
    return [uncheckable(rule, "Verify abbreviations are used consistently throughout headings, tables and text.")];
  }
  if (params.rule === "authorized-only" || params.rule === "authorized-titles-only") {
    return [uncheckable(rule, "Authorized-abbreviation list check: requires the Ch 16 Annex reference data to be loaded.")];
  }
  if (params.rule === "no-introduction-in-headings") {
    return [uncheckable(rule, "Verify abbreviations are not introduced in headings.")];
  }
  if (params.allowed === "none") {
    const acro = text.match(/\b[A-Z]{2,}\b/g) || [];
    if (acro.length) return [finding(rule, `Abbreviations must not be used in letters to civilians (found: ${[...new Set(acro)].slice(0, 6).join(", ")}).`)];
    return [];
  }
  if (params.allowed === "all") return [];
  return [];
};

module.exports = { EVALUATORS, finding, uncheckable };
