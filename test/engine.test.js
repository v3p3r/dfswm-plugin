"use strict";

const test = require("node:test");
const assert = require("node:assert");

const { runCompliance, ruleApplies, groupBySeverity } = require("../src/engine/engine.js");
const { EVALUATORS } = require("../src/engine/evaluators.js");
const { normalizeModel } = require("../src/engine/model.js");
const { para, model, servicePaperModel } = require("./helpers/mock-model.js");

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

function find(report, ruleId) {
  return report.findings.filter((f) => f.ruleId === ruleId);
}

/* ------------------------------------------------------------------ */

test("ruleApplies: general applies to everything; specific types filter", () => {
  assert.ok(ruleApplies({ appliesTo: ["general"] }, "signal"));
  assert.ok(ruleApplies({ appliesTo: ["signal"] }, "signal"));
  assert.ok(!ruleApplies({ appliesTo: ["signal"] }, "service-paper"));
});

test("font-name: flags disallowed fonts", () => {
  const m = model([para("Hello world", { fontName: "Arial" }), para("Fine text", { fontName: "Times New Roman" })]);
  const report = runCompliance(m, {
    docType: "general",
    rulesets: [{ id: "core", title: "Core", rules: [rule("core.font", "error", { type: "font-name", scope: "all", params: { allowed: ["Times New Roman"] } })] }],
  });
  assert.strictEqual(find(report, "core.font").length, 1);
  assert.strictEqual(find(report, "core.font")[0].fix.action, "setFont");
});

test("font-size: flags wrong body size", () => {
  const m = model([para("Body", { fontSize: 11 })]);
  const report = runCompliance(m, {
    rulesets: [{ id: "core", rules: [rule("size", "error", { type: "font-size", scope: "body", params: { size: 12, comparison: "eq" } })] }],
  });
  assert.strictEqual(find(report, "size").length, 1);
  assert.match(find(report, "size")[0].message, /not 12 pt/);
});

test("justification: flags left-aligned body text", () => {
  const m = model([para("Left aligned", { alignment: "left" })]);
  const report = runCompliance(m, {
    rulesets: [{ id: "core", rules: [rule("just", "error", { type: "justification", scope: "body", params: { alignment: "justified" } })] }],
  });
  assert.strictEqual(find(report, "just").length, 1);
});

test("margin: flags 2.5cm left margin when 2cm required", () => {
  const m = model([para("x")], {
    sections: [{ margins: { left: 2.5, right: 2, top: 2, bottom: 2 }, headerDistanceCm: 1.5, footerDistanceCm: 1.5, headers: { primary: [] }, footers: { primary: [] } }],
  });
  const report = runCompliance(m, {
    rulesets: [{ id: "core", rules: [rule("m", "error", { type: "margin", scope: "all", params: { sides: ["left", "right"], valueCm: 2 } })] }],
  });
  assert.strictEqual(find(report, "m").length, 1);
  assert.match(find(report, "m")[0].message, /left margin is 2.50 cm/);
});

test("regex present/absent semantics", () => {
  const ok = runCompliance(model([para("We move at 0001 hours tonight.")]), {
    rulesets: [{ id: "dt", rules: [rule("no-midnight", "error", { type: "regex", scope: "body", params: { pattern: "\\bmidnight\\b", expect: "absent" } })] }],
  });
  assert.strictEqual(find(ok, "no-midnight").length, 0);

  const bad = runCompliance(model([para("We move at midnight tonight.")]), {
    rulesets: [{ id: "dt", rules: [rule("no-midnight", "error", { type: "regex", scope: "body", params: { pattern: "\\bmidnight\\b", expect: "absent" } })] }],
  });
  assert.strictEqual(find(bad, "no-midnight").length, 1);
});

test("element-order: service paper mandatory headings in order", () => {
  const m = servicePaperModel();
  const report = runCompliance(m, {
    docType: "service-paper",
    rulesets: [{
      id: "sp",
      rules: [rule("sp.mandatory", "error", {
        type: "element-order",
        scope: "body",
        params: { rule: "mandatory-sequence", mandatory: ["INTRODUCTION", "AIM", "CONCLUSION", "RECOMMENDATIONS"], order: true },
      })],
    }],
  });
  assert.strictEqual(find(report, "sp.mandatory").length, 0);

  // Reorder: RECOMMENDATIONS before CONCLUSION
  const bad = model([
    para("INTRODUCTION", { bold: true, headingLevel: "main" }),
    para("AIM", { bold: true, headingLevel: "main" }),
    para("RECOMMENDATIONS", { bold: true, headingLevel: "main" }),
    para("CONCLUSION", { bold: true, headingLevel: "main" }),
  ]);
  const report2 = runCompliance(bad, {
    docType: "service-paper",
    rulesets: [{
      id: "sp",
      rules: [rule("sp.mandatory", "error", {
        type: "element-order", scope: "body",
        params: { rule: "mandatory-sequence", mandatory: ["INTRODUCTION", "AIM", "CONCLUSION", "RECOMMENDATIONS"], order: true },
      })],
    }],
  });
  assert.strictEqual(find(report2, "sp.mandatory").length, 1);
  assert.match(find(report2, "sp.mandatory")[0].message, /out of order/);
});

test("no-single-subparagraph: paragraph with exactly one sub-paragraph", () => {
  const m = model([
    para("1.\tFirst paragraph"),
    para("a.\tOnly sub-paragraph"),
  ]);
  const report = runCompliance(m, {
    rulesets: [{ id: "num", rules: [rule("num.single", "error", { type: "no-single-subparagraph", scope: "body", params: {} })] }],
  });
  assert.strictEqual(find(report, "num.single").length, 1);
});

test("manual rules render as manual checklist items", () => {
  const m = model([para("x")]);
  const report = runCompliance(m, {
    rulesets: [{ id: "sec", rules: [rule("sec.precedence-authority", "info", { type: "precedence-marking", scope: "first-page", params: {} }, { manual: true })] }],
  });
  const f = find(report, "sec.precedence-authority");
  assert.strictEqual(f.length, 1);
  assert.strictEqual(f[0].manual, true);
  assert.strictEqual(report.counts.manual, 1);
});

test("appliesTo filters rules by document type", () => {
  const m = model([para("x")]);
  const ruleset = { id: "corr", rules: [rule("corr.formal", "error", { type: "regex", scope: "body", params: { pattern: "^Sir\\b", expect: "present" } }, { appliesTo: ["formal-letter"] })] };
  const letter = runCompliance(m, { docType: "formal-letter", rulesets: [ruleset] });
  assert.strictEqual(find(letter, "corr.formal").length, 1);
  const paper = runCompliance(m, { docType: "service-paper", rulesets: [ruleset] });
  assert.strictEqual(find(paper, "corr.formal").length, 0);
});

test("unknown check type produces an uncheckable finding, not a crash", () => {
  const m = model([para("x")]);
  const report = runCompliance(m, {
    rulesets: [{ id: "x", rules: [rule("x.unknown", "error", { type: "made-up-type", scope: "body", params: {} })] }],
  });
  assert.strictEqual(find(report, "x.unknown")[0].checkable, false);
});

test("groupBySeverity buckets correctly", () => {
  const report = {
    findings: [
      { ruleId: "a", severity: "error" },
      { ruleId: "b", severity: "warning" },
      { ruleId: "c", severity: "info" },
      { ruleId: "d", manual: true },
    ],
  };
  const { groups } = groupBySeverity(report.findings);
  assert.strictEqual(groups.error.length, 1);
  assert.strictEqual(groups.warning.length, 1);
  assert.strictEqual(groups.info.length, 1);
  assert.strictEqual(groups.manual.length, 1);
});

test("normalizeModel is idempotent and adds indices", () => {
  const m = normalizeModel({ docType: "signal", paragraphs: [{ text: "ALL CAPS", allCaps: true }] });
  assert.strictEqual(m.paragraphs[0].index, 0);
  assert.strictEqual(m.paragraphs[0].fontName, null);
  assert.strictEqual(m.docType, "signal");
  assert.strictEqual(m.bodyText, "ALL CAPS");
});

test("EVALUATORS registry covers the check types used by the rulesets", async () => {
  const fs = require("node:fs");
  const path = require("node:path");
  const rulesetsDir = path.resolve(__dirname, "..", "..", "dfswm-rulesets", "rulesets");
  const files = fs.readdirSync(rulesetsDir).filter((f) => f.endsWith(".json"));
  const used = new Set();
  for (const f of files) {
    const rs = JSON.parse(fs.readFileSync(path.join(rulesetsDir, f), "utf8"));
    for (const r of rs.rules) {
      if (r.check && r.check.type) used.add(r.check.type);
    }
  }
  const implemented = new Set(Object.keys(EVALUATORS));
  const missing = [...used].filter((t) => !implemented.has(t));
  assert.deepStrictEqual(missing, [], `check types without an evaluator: ${missing.join(", ")}`);
});
