/**
 * DFSWM Compliance Plugin — engine.
 *
 * runCompliance(model, { docType, rulesets, references }) => Report
 *
 *   model      raw OR normalised document model (see model.js); it is
 *              normalised once inside runCompliance before any evaluator runs
 *   docType    selected document type id (see dfswm-rulesets/index.json)
 *   rulesets   array of ruleset objects (id, rules[])
 *   references optional { abbreviations: string[] } for abbreviation checks
 *
 * Report:
 *   { docType, counts: { error, warning, info, manual },
 *     findings: Finding[], rulesEvaluated, rulesSkipped }
 */
"use strict";

const { EVALUATORS, finding, uncheckable } = require("./evaluators.js");
const { normalizeModel, textForScope, fullText } = require("./model.js");

/**
 * True when a rule applies to the selected document type.
 * `general` applies to everything; otherwise the doc type must be listed.
 */
function ruleApplies(rule, docType) {
  const applies = rule.appliesTo || ["general"];
  return applies.includes("general") || applies.includes(docType);
}

function runRule(rule, model, ctx) {
  // Manual rules are always rendered as human checklist items.
  if (rule.manual) {
    return [{ ...uncheckable(rule, rule.notes || rule.description), manual: true }];
  }
  const evaluator = EVALUATORS[rule.check && rule.check.type];
  if (!evaluator) {
    return [{ ...uncheckable(rule, `No evaluator implemented for check type '${rule.check && rule.check.type}'.`), checkable: false }];
  }
  try {
    const results = evaluator(model, rule, ctx) || [];
    return Array.isArray(results) ? results : [results];
  } catch (err) {
    return [{ ...uncheckable(rule, `Evaluator error: ${err.message}`), checkable: false }];
  }
}

/**
 * Run one or more rulesets against the model.
 * @returns {object} report
 */
function runCompliance(rawModel, { docType = "general", rulesets = [], references = null } = {}) {
  // Normalise once so every evaluator sees bodyText, indices and typed
  // fields regardless of whether the caller passed a raw extractor model
  // (Word/PowerPoint) or an already-normalised model. Idempotent.
  const model = normalizeModel(rawModel);
  const ctx = { references, textForScope, fullText };
  const findings = [];
  let rulesEvaluated = 0;
  let rulesSkipped = 0;

  for (const ruleset of rulesets) {
    for (const rule of ruleset.rules || []) {
      if (!ruleApplies(rule, docType)) {
        rulesSkipped += 1;
        continue;
      }
      rulesEvaluated += 1;
      const results = runRule(rule, model, ctx);
      for (const f of results) {
        // Carry the finding's locations onto its fix object so the host
        // fixer can target exactly the flagged paragraphs instead of the
        // whole document. Document-level fixes (margins, page numbers,
        // copy numbers) get an empty list and are applied document-wide.
        if (f.fix && f.fix.locations === undefined) {
          f.fix = { ...f.fix, locations: f.locations || [] };
        }
        findings.push({ ...f, rulesetId: ruleset.id, rulesetTitle: ruleset.title || ruleset.id });
      }
    }
  }

  const counts = { error: 0, warning: 0, info: 0, manual: 0 };
  for (const f of findings) {
    if (f.manual) counts.manual += 1;
    else if (f.severity === "error") counts.error += 1;
    else if (f.severity === "warning") counts.warning += 1;
    else counts.info += 1;
  }

  return { docType, counts, findings, rulesEvaluated, rulesSkipped };
}

/**
 * Group findings by severity for the UI.
 */
function groupBySeverity(findings) {
  const order = ["error", "warning", "info", "manual"];
  const groups = { error: [], warning: [], info: [], manual: [] };
  for (const f of findings) {
    const key = f.manual ? "manual" : order.includes(f.severity) ? f.severity : "info";
    groups[key].push(f);
  }
  return { order, groups };
}

module.exports = { runCompliance, runRule, ruleApplies, groupBySeverity };
