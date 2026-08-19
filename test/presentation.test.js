"use strict";

const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");

const { runCompliance } = require("../src/engine/engine.js");
const { normalizeModel } = require("../src/engine/model.js");
const { mapPresentationToModel } = require("../src/powerpoint/extractor.js");

function slide(name, shapes) {
  return { name, shapes };
}

function shape(name, paragraphs, hasTextFrame = true) {
  return { name, hasTextFrame, paragraphs };
}

function p(text, over = {}) {
  return {
    text,
    fontName: over.fontName ?? "Arial",
    fontSize: over.fontSize ?? 24,
    bold: !!over.bold,
    italic: !!over.italic,
    underline: !!over.underline,
    alignment: over.alignment ?? "Left",
  };
}

/* ------------------------------------------------------------------ */

test("mapPresentationToModel: tags slideIndex, maps alignment, keeps text", () => {
  const m = mapPresentationToModel([
    slide("Slide 1", [
      shape("Title", [p("OPERATION IRON SHIELD", { bold: true, alignment: "Center" })]),
      shape("Body", [p("The plan for tonight.", { alignment: "Left" })]),
    ]),
    slide("Slide 2", [shape("Body", [p("RESTRICTED", { alignment: "Left" })])]),
  ]);

  assert.strictEqual(m.docType, "presentation");
  assert.strictEqual(m.paragraphs.length, 3);
  assert.strictEqual(m.paragraphs[0].slideIndex, 0);
  assert.strictEqual(m.paragraphs[1].slideIndex, 0);
  assert.strictEqual(m.paragraphs[2].slideIndex, 1);
  assert.strictEqual(m.paragraphs[0].alignment, "center");
  assert.strictEqual(m.paragraphs[1].alignment, "left");
  assert.strictEqual(m.paragraphs[0].text, "OPERATION IRON SHIELD");
  assert.strictEqual(m.paragraphs[0].bold, true);
  // headings heuristic: centered + all-caps short text => subject
  assert.strictEqual(m.paragraphs[0].headingLevel, "subject");
  // sections/tables/footnotes are empty for slides
  assert.deepStrictEqual(m.sections, []);
  assert.deepStrictEqual(m.tables, []);
  assert.deepStrictEqual(m.footnotes, []);
});

test("mapPresentationToModel: skips shapes without text frames and blank paragraphs", () => {
  const m = mapPresentationToModel([
    slide("S1", [
      shape("Picture", [], false),
      shape("Text", [p("   "), p("Real text")]),
    ]),
  ]);
  assert.strictEqual(m.paragraphs.length, 1);
  assert.strictEqual(m.paragraphs[0].text, "Real text");
});

test("normalizeModel passes slideIndex through", () => {
  const m = normalizeModel({ paragraphs: [{ text: "hi", slideIndex: 3 }] });
  assert.strictEqual(m.paragraphs[0].slideIndex, 3);
  assert.strictEqual(m.paragraphs[0].index, 0);
});

test("presentation ruleset runs against a slide model: finds violations", () => {
  const ruleset = JSON.parse(
    fs.readFileSync(path.resolve(__dirname, "..", "..", "dfswm-rulesets", "rulesets", "presentation.json"), "utf8"),
  );

  const m = mapPresentationToModel([
    slide("Slide 1", [
      // Comic Sans is not an approved font -> pres.font-family
      shape("Title", [p("BRIEFING TITLE", { fontName: "Comic Sans MS", bold: true, alignment: "Center" })]),
      // underlined text -> pres.no-underline
      shape("Body", [p("This is underlined.", { underline: true })]),
      // no security marking anywhere -> pres.security-marking
      shape("Body", [p("The aim of this briefing is to inform.")]),
    ]),
  ]);

  const report = runCompliance(m, { docType: "presentation", rulesets: [ruleset] });
  const byRule = {};
  for (const f of report.findings) byRule[f.ruleId] = f;

  assert.ok(byRule["pres.font-family"], "font-family rule should fire");
  assert.match(byRule["pres.font-family"].message, /Comic Sans MS/);
  assert.ok(byRule["pres.no-underline"], "underline rule should fire");
  assert.ok(byRule["pres.security-marking"], "security marking rule should fire");

  // manual rules always render as checklist items
  for (const id of ["pres.title-slide", "pres.concise-essentials", "pres.aim-stated", "pres.legible-sizes", "pres.slide-numbers"]) {
    assert.ok(byRule[id], `manual rule ${id} should be present`);
    assert.strictEqual(byRule[id].manual, true);
  }
});

test("presentation ruleset passes a compliant deck", () => {
  const ruleset = JSON.parse(
    fs.readFileSync(path.resolve(__dirname, "..", "..", "dfswm-rulesets", "rulesets", "presentation.json"), "utf8"),
  );

  const m = mapPresentationToModel([
    slide("Slide 1", [
      shape("Title", [p("Briefing Title", { fontName: "Arial", bold: true, alignment: "Center" })]),
      // The security marking is explained on first use, so the
      // abbreviation-usage rule does not flag it.
      shape("Body", [p("Security classification: Restricted (RESTRICTED).", { fontName: "Arial", alignment: "Left" })]),
      shape("Body", [p("The aim of this briefing is to inform.", { fontName: "Arial", alignment: "Left" })]),
    ]),
    slide("Slide 2", [
      shape("Body", [p("RESTRICTED", { fontName: "Arial", alignment: "Left" })]),
      shape("Body", [p("Move at 13 Aug 26.", { fontName: "Arial", alignment: "Left" })]),
    ]),
  ]);

  const report = runCompliance(m, { docType: "presentation", rulesets: [ruleset] });

  // none of the checkable rules should fire
  for (const f of report.findings) {
    if (!f.manual) {
      assert.fail(`unexpected finding for ${f.ruleId}: ${f.message}`);
    }
  }
  // all manual checklist items still present
  assert.strictEqual(report.counts.manual, 5);
});

test("presentation ruleset does not apply to Word document types", () => {
  const ruleset = JSON.parse(
    fs.readFileSync(path.resolve(__dirname, "..", "..", "dfswm-rulesets", "rulesets", "presentation.json"), "utf8"),
  );
  const m = mapPresentationToModel([
    slide("S1", [shape("Body", [p("Comic Sans text", { fontName: "Comic Sans MS" })])]),
  ]);
  const report = runCompliance(m, { docType: "brief", rulesets: [ruleset] });
  assert.strictEqual(report.rulesEvaluated, 0);
  assert.strictEqual(report.findings.length, 0);
});
