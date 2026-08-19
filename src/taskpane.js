/**
 * DFSWM Compliance Plugin — task pane controller.
 * Wires the UI to the ruleset loader, engine and the host-specific
 * extractor/fixer (Word or PowerPoint).
 */
"use strict";

(function () {
  const { runCompliance, groupBySeverity } = window.DFSWM;

  // Rulesets are synced into dfswm-plugin/rulesets/ and dfswm-plugin/index.json
  // by scripts/sync-rulesets.js (paths relative to src/taskpane.html).
  const RULESET_BASE = "../rulesets";
  const INDEX_URL = "../index.json";

  let host = null; // "Word" | "PowerPoint"
  let extractModel = null; // host-specific extractor
  let applyFix = null; // host-specific fixer
  let index = null;
  let rulesets = [];
  let docType = "general";
  let activeRulesetIds = null; // null => all for the doc type
  let lastReport = null;
  let lastModel = null;

  const $ = (id) => document.getElementById(id);

  function setStatus(text, spin) {
    const el = $("status");
    el.textContent = text;
    el.className = "status" + (spin ? " spin" : "");
  }

  async function loadJson(url) {
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error(`Failed to load ${url} (${res.status})`);
    return res.json();
  }

  async function loadAll() {
    setStatus("Loading rulesets…", true);
    try {
      index = await loadJson(INDEX_URL);
      const files = new Set();
      for (const cfg of Object.values(index.documentTypes)) {
        for (const rs of cfg.rulesets) files.add(rs);
      }
      rulesets = [];
      for (const rsId of files) {
        const rel = index.rulesetFiles[rsId];
        if (!rel) continue;
        rulesets.push(await loadJson(`${RULESET_BASE}/${rel.split("/").pop()}`));
      }
      renderDocTypes();
      renderRulesetChips();
      setStatus("Ready");
      $("load-rulesets").hidden = true;
    } catch (err) {
      setStatus(`Error: ${err.message}`);
      $("load-rulesets").hidden = false;
    }
  }

  function renderDocTypes() {
    const sel = $("doc-type");
    sel.innerHTML = "";
    const keys = Object.keys(index.documentTypes).sort();
    for (const key of keys) {
      // PowerPoint only supports the presentation doc type.
      if (host === "PowerPoint" && key !== "presentation") continue;
      const opt = document.createElement("option");
      opt.value = key;
      opt.textContent = index.documentTypes[key].label;
      sel.appendChild(opt);
    }
    if (host === "PowerPoint") docType = "presentation";
    sel.value = docType;
    sel.addEventListener("change", () => {
      docType = sel.value;
      activeRulesetIds = null;
      renderRulesetChips();
    });
  }

  function rulesetsForDocType() {
    const cfg = index && index.documentTypes[docType];
    const ids = cfg ? cfg.rulesets : [];
    return rulesets.filter((r) => ids.includes(r.id));
  }

  function renderRulesetChips() {
    const wrap = $("ruleset-select");
    wrap.innerHTML = "";
    const available = rulesetsForDocType();
    if (!available.length) {
      wrap.innerHTML = '<span class="hint">No rulesets for this document type.</span>';
      return;
    }
    for (const rs of available) {
      const chip = document.createElement("span");
      chip.className = "chip active";
      chip.textContent = rs.title.split("—")[0].split("(")[0].trim() || rs.id;
      chip.dataset.id = rs.id;
      chip.addEventListener("click", () => {
        chip.classList.toggle("active");
        const active = wrap.querySelectorAll(".chip.active");
        activeRulesetIds = active.length === available.length ? null : Array.from(active).map((c) => c.dataset.id);
      });
      wrap.appendChild(chip);
    }
  }

  function selectedRulesets() {
    let list = rulesetsForDocType();
    if (activeRulesetIds) list = list.filter((r) => activeRulesetIds.includes(r.id));
    return list;
  }

  function esc(s) {
    const div = document.createElement("div");
    div.textContent = String(s == null ? "" : s);
    return div.innerHTML;
  }

  function renderSummary(report) {
    const c = report.counts;
    const items = [
      { cls: "error", num: c.error, lbl: "Errors" },
      { cls: "warning", num: c.warning, lbl: "Warnings" },
      { cls: "info", num: c.info, lbl: "Info" },
      { cls: "manual", num: c.manual, lbl: "Manual" },
    ];
    $("summary").innerHTML = items
      .map((i) => `<div class="sum-item ${i.cls}"><div class="num">${i.num}</div><div class="lbl">${i.lbl}</div></div>`)
      .join("");
  }

  function renderFinding(f) {
    const el = document.createElement("div");
    const cls = f.manual ? "manual" : ["error", "warning", "info"].includes(f.severity) ? f.severity : "info";
    el.className = `finding ${cls}`;
    const loc = f.locations && f.locations.length
      ? `<div class="finding-loc">${esc(f.locations[0].snippet || f.locations[0].paragraphIndex)}</div>`
      : "";
    const ref = f.chapter ? `Ch ${f.chapter}${f.paragraphs ? ` ¶${f.paragraphs}` : ""}` : "";
    let actions = "";
    if (f.fix && f.fix.action) {
      actions = `<button class="btn-fix" data-fix-index="${f._fixIndex}">Fix</button>`;
    }
    el.innerHTML = `
      <div class="finding-head">
        <span class="finding-title">${esc(f.title)}</span>
        <span class="finding-ref">${esc(ref)}</span>
      </div>
      <div class="finding-msg">${esc(f.message)}</div>
      ${loc}
      <div class="finding-actions">${actions}</div>`;
    return el;
  }

  function renderResults(report) {
    const wrap = $("results");
    wrap.innerHTML = "";
    const { groups, order } = groupBySeverity(report.findings);
    let count = 0;
    const labels = { error: "Errors", warning: "Warnings", info: "Information", manual: "Needs manual review" };
    for (const key of order) {
      const list = groups[key];
      if (!list.length) continue;
      const group = document.createElement("div");
      group.className = "find-group";
      group.innerHTML = `<div class="group-head">${labels[key]} (${list.length})</div>`;
      list.forEach((f) => {
        f._fixIndex = count++;
        group.appendChild(renderFinding(f));
      });
      wrap.appendChild(group);
    }
    if (!count) wrap.innerHTML = '<div class="empty">✅ No compliance issues found.</div>';
  }

  async function runCheck() {
    setStatus("Checking document…", true);
    $("run-check").disabled = true;
    try {
      lastModel = await extractModel();
      lastModel.docType = docType;
      const report = runCompliance(lastModel, { docType, rulesets: selectedRulesets() });
      lastReport = report;
      $("config-pane").hidden = true;
      $("results-pane").hidden = false;
      renderSummary(report);
      renderResults(report);
      setStatus(`${report.findings.length} finding(s)`);
    } catch (err) {
      setStatus(`Check failed: ${err.message}`);
    } finally {
      $("run-check").disabled = false;
    }
  }

  async function onFixClick(e) {
    const btn = e.target.closest(".btn-fix");
    if (!btn || !lastReport) return;
    const idx = Number(btn.dataset.fixIndex);
    const findingObj = lastReport.findings.find((f) => f._fixIndex === idx);
    if (!findingObj || !findingObj.fix) return;
    setStatus("Applying fix…", true);
    const result = await applyFix(findingObj.fix);
    setStatus(result.ok ? `Fix applied: ${result.message}` : `Fix failed: ${result.message}`);
    if (result.ok) setTimeout(runCheck, 300);
  }

  function bind() {
    $("run-check").addEventListener("click", runCheck);
    $("back").addEventListener("click", () => {
      $("results-pane").hidden = true;
      $("config-pane").hidden = false;
    });
    $("load-rulesets").addEventListener("click", loadAll);
    $("results").addEventListener("click", onFixClick);
  }

  Office.onReady((info) => {
    host = info.host;
    if (host === "Word") {
      extractModel = window.DFSWM.extractDocumentModel;
      applyFix = window.DFSWM.applyFix;
    } else if (host === "PowerPoint") {
      extractModel = window.DFSWM.extractPresentationModel;
      applyFix = window.DFSWM.applyPptFix;
    } else {
      setStatus("This add-in supports Microsoft Word and PowerPoint.");
      return;
    }
    bind();
    loadAll();
  });
})();
