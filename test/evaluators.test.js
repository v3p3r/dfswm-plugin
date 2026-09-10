"use strict";

const test = require("node:test");
const assert = require("node:assert");

const { runCompliance } = require("../src/engine/engine.js");
const { normalizeModel, textForScope, fullText } = require("../src/engine/model.js");
const { para, model } = require("./helpers/mock-model.js");

function rule(id, severity, check, extra = {}) {
  return {
    id,
    title: `Rule ${id}`,
    description: "",
    chapter: "2",
    paragraphs: "204",
    severity,
    appliesTo: extra.appliesTo || ["general"],
    check,
    ...extra,
  };
}

function single(report, ruleId) {
  return report.findings.find((f) => f.ruleId === ruleId);
}

/* ------------------- fix location wiring (engine) ------------------- */

test("engine attaches finding locations onto paragraph-level fixes", () => {
  const m = model([
    para("Good size", { fontSize: 12 }),
    para("Wrong size", { fontSize: 10 }),
    para("Also wrong", { fontSize: 9 }),
  ]);
  const report = runCompliance(m, {
    rulesets: [{
      id: "core",
      rules: [rule("size", "error", { type: "font-size", scope: "body", params: { size: 12 } })],
    }],
  });
  const f = single(report, "size");
  assert.ok(f.fix);
  assert.deepStrictEqual(f.fix.locations.map((l) => l.paragraphIndex), [1, 2]);
});

test("document-level fixes (margins) carry an empty location list", () => {
  const m = model([para("x")], {
    sections: [{ margins: { left: 3, right: 2, top: 2, bottom: 2 } }],
  });
  const report = runCompliance(m, {
    rulesets: [{
      id: "core",
      rules: [rule("m", "error", { type: "margin", scope: "all", params: { sides: ["left"], valueCm: 2 } })],
    }],
  });
  const f = single(report, "m");
  assert.ok(f.fix);
  assert.deepStrictEqual(f.fix.locations, []);
});

test("manual and uncheckable findings never carry a fix", () => {
  const report = runCompliance(model([para("x")]), {
    rulesets: [{
      id: "x",
      rules: [
        rule("manual", "info", { type: "regex", scope: "body", params: {} }, { manual: true }),
        rule("unknown", "error", { type: "made-up-type", scope: "body", params: {} }),
      ],
    }],
  });
  assert.strictEqual(single(report, "manual").fix, null);
  assert.strictEqual(single(report, "unknown").fix, null);
});

test("each fixable finding's locations are a subset of the findings' own locations", () => {
  const m = model([
    para("One", { alignment: "left" }),
    para("Two", { alignment: "left" }),
    para("Three", { alignment: "justified" }),
  ]);
  const report = runCompliance(m, {
    rulesets: [{
      id: "core",
      rules: [rule("just", "error", { type: "justification", scope: "body", params: { alignment: "justified" } })],
    }],
  });
  const f = single(report, "just");
  assert.deepStrictEqual(f.fix.locations, f.locations);
  assert.deepStrictEqual(f.locations.map((l) => l.paragraphIndex), [0, 1]);
});

/* ------------------- applyRegexFix replacements ------------------- */

test("midnight autofix uses the '0000 hours' replacement", () => {
  const report = runCompliance(model([para("We move at midnight.")]), {
    rulesets: [{
      id: "dt",
      rules: [rule("midnight", "error", {
        type: "regex", scope: "body",
        params: { pattern: "\\bmidnight\\b", expect: "absent" },
      }, { autofix: { available: true } })],
    }],
  });
  const f = single(report, "midnight");
  assert.strictEqual(f.fix.action, "applyRegexFix");
  assert.deepStrictEqual(f.fix.params, { pattern: "\\bmidnight\\b", replacement: "0000 hours" });
});

test("punctuation-spacing autofix collapses space runs, keeping the punctuation", () => {
  const report = runCompliance(model([para("One.  Two.")]), {
    rulesets: [{
      id: "core",
      rules: [rule("spacing", "error", {
        type: "regex", scope: "body",
        params: { pattern: "[,;:.!?] {2,}", expect: "absent" },
      }, { autofix: { available: true } })],
    }],
  });
  const f = single(report, "spacing");
  assert.strictEqual(f.fix.action, "applyRegexFix");
  assert.deepStrictEqual(f.fix.params, { pattern: "([,;:.!?]) {2,}", replacement: "$1 " });
});

test("abbreviation full-stop autofix only targets mid-sentence dots", () => {
  const report = runCompliance(model([para("The JCSC. training year.")]), {
    rulesets: [{
      id: "abb",
      rules: [rule("nopunct", "error", {
        type: "regex", scope: "body",
        params: { pattern: "\\b[A-Z]{2,}\\.", expect: "absent" },
      }, { autofix: { available: true } })],
    }],
  });
  const f = single(report, "nopunct");
  assert.strictEqual(f.fix.action, "applyRegexFix");
  // Sentence-final dots must never be stripped: the fix pattern requires a
  // lower-case/digit word after the dot.
  assert.strictEqual(f.fix.params.pattern, "\\b([A-Z]{2,})\\.(?=\\s+[a-z0-9])");
  assert.strictEqual(f.fix.params.replacement, "$1");
});

test("unrecognised pattern with autofix emits no fix (never blind-delete)", () => {
  const report = runCompliance(model([para("Some foo here.")]), {
    rulesets: [{
      id: "x",
      rules: [rule("foo", "error", {
        type: "regex", scope: "body",
        params: { pattern: "\\bfoo\\b", expect: "absent" },
      }, { autofix: { available: true } })],
    }],
  });
  const f = single(report, "foo");
  assert.strictEqual(f.fix, null);
});

test("rule-declared autofix.replacement is used verbatim", () => {
  const report = runCompliance(model([para("Something to fix.")]), {
    rulesets: [{
      id: "x",
      rules: [rule("custom", "error", {
        type: "regex", scope: "body",
        params: { pattern: "\\bto fix\\b", expect: "absent" },
      }, { autofix: { available: true, replacement: "fixed" } })],
    }],
  });
  const f = single(report, "custom");
  assert.strictEqual(f.fix.action, "applyRegexFix");
  assert.deepStrictEqual(f.fix.params, { pattern: "\\bto fix\\b", replacement: "fixed" });
});

/* ------------------- heading-style fix ------------------- */

test("heading-style fix carries the computed size and target locations", () => {
  const m = model([
    para("Body body body body body body", { fontSize: 12 }),
    para("MAIN HEADING", { bold: true, allCaps: true, alignment: "center", fontSize: 11, headingLevel: "main" }),
  ]);
  const report = runCompliance(m, {
    rulesets: [{
      id: "hdg",
      rules: [rule("main", "error", {
        type: "heading-style", scope: "main-heading",
        params: { alignment: "center", bold: true, caps: true, sizeOffsetPoints: 2 },
      })],
    }],
  });
  const f = single(report, "main");
  assert.ok(f);
  assert.strictEqual(f.fix.action, "fixHeading");
  assert.strictEqual(f.fix.params.size, 14);
  assert.deepStrictEqual(f.fix.locations.map((l) => l.paragraphIndex), [1]);
});

/* ------------------- other evaluator behaviour ------------------- */

test("italic evaluator emits a setItalic fix for quoted paragraphs", () => {
  const report = runCompliance(model([para('He said "carry on" and left.')]), {
    rulesets: [{
      id: "core",
      rules: [rule("it", "warning", { type: "italic", scope: "body", params: { expected: true } })],
    }],
  });
  const f = single(report, "it");
  assert.ok(f);
  assert.strictEqual(f.fix.action, "setItalic");
  assert.deepStrictEqual(f.fix.params, { italic: true });
});

test("keep-with-next fix carries the keep flags for signature blocks", () => {
  const m = model([
    para("Some body text."),
    para("V OJWANG", { bold: true }),
    para("Lt Col"),
  ]);
  const report = runCompliance(m, {
    rulesets: [{
      id: "core",
      rules: [rule("sig", "error", {
        type: "keep-with-next", scope: "signature-block",
        params: { keepWithPrevious: true },
      })],
    }],
  });
  const f = single(report, "sig");
  assert.ok(f);
  assert.deepStrictEqual(f.fix.params, { keepWithPrevious: true });
});

test("required-element flags missing literal phrases", () => {
  const report = runCompliance(model([para("No reference list here.")]), {
    rulesets: [{
      id: "corr",
      rules: [rule("ref", "error", {
        type: "required-element", scope: "body",
        params: { elements: ["list of references"] },
      })],
    }],
  });
  const f = single(report, "ref");
  assert.ok(f);
  assert.match(f.message, /Missing required element/);
});

test("date-format present/absent semantics", () => {
  const params = { pattern: "^\\d{1,2} (January|February|March|April|May|June|July|August|September|October|November|December) \\d{4}$" };
  const good = runCompliance(model([para("1 July 2024")]), {
    rulesets: [{ id: "dt", rules: [rule("date", "error", { type: "date-format", scope: "body", params })] }],
  });
  assert.strictEqual(single(good, "date"), undefined);

  const bad = runCompliance(model([para("July 1, 2024")]), {
    rulesets: [{ id: "dt", rules: [rule("date", "error", { type: "date-format", scope: "body", params })] }],
  });
  assert.ok(single(bad, "date"));
});

test("paragraph-numbering flags a marker that mismatches its indent level", () => {
  const m = model([
    para("1.\tFirst level", { blockIndentCm: 0 }),
    para("(1)\tWrong marker for level two", { blockIndentCm: 1.0 }),
  ]);
  const report = runCompliance(m, {
    rulesets: [{
      id: "num",
      rules: [rule("hier", "error", { type: "paragraph-numbering", scope: "body", params: {} })],
    }],
  });
  const f = single(report, "hier");
  assert.ok(f);
  assert.match(f.message, /numbering scheme/);
});

test("security-marking requires a recognised classification", () => {
  const report = runCompliance(model([para("Nothing classified here.")]), {
    rulesets: [{
      id: "sec",
      rules: [rule("class", "error", {
        type: "security-marking", scope: "all",
        params: { mustMatchDocumentClassification: true },
      })],
    }],
  });
  const f = single(report, "class");
  assert.ok(f);
  assert.match(f.message, /classification/);
});

/* ------------------- model normalisation ------------------- */

test("normalizeModel parses tables, footnotes and header/footer scopes", () => {
  const m = normalizeModel({
    docType: "general",
    paragraphs: [{ text: "Body" }],
    sections: [{
      margins: null,
      headers: { primary: ["CLASSIFICATION"], firstPage: [], evenPages: [] },
      footers: { primary: ["1"], firstPage: [], evenPages: [] },
    }],
    tables: [{ headerRow: [{ text: "Col A" }], rows: [[{ text: "x", fontSize: 10 }]] }],
    footnotes: [{ text: "a note", fontSize: 8 }],
  });
  assert.strictEqual(m.tables[0].headerRow[0], "Col A");
  assert.strictEqual(m.tables[0].rows[0][0].text, "x");
  assert.strictEqual(m.footnotes[0].fontSize, 8);
  assert.match(textForScope(m, "header"), /CLASSIFICATION/);
  assert.match(textForScope(m, "table"), /Col A/);
  assert.match(textForScope(m, "footnote"), /a note/);
  assert.match(fullText(m), /CLASSIFICATION/);
  assert.match(fullText(m), /x/);
});