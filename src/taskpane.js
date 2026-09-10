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

  const SETTINGS_KEY_DOC_TYPE = "dfswm.docType";
  const SETTINGS_KEY_RULESETS = "dfswm.rulesets";

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

  /* ------------------------- settings persistence ------------------------- */
  // Preference lives in Office.settings when running inside Word/PowerPoint
  // so it travels with the document; localStorage is the fallback for plain
  // browser testing of the task pane.

  function hostSettings() {
    try {
      return Office.context && Office.context.document && Office.context.document.settings;
    } catch (e) {
      return null;
    }
  }

  function saveSetting(key, value) {
    try {
      const s = hostSettings();
      if (s) {
        s.set(key, value);
        s.saveAsync(() => {});
        return;
      }
    } catch (e) { /* fall through to localStorage */ }
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch (e) { /* ignore */ }
  }

  function readSetting(key) {
    try {
      const s = hostSettings();
      if (s) return s.get(key);
    } catch (e) { /* fall through to localStorage */ }
    try {
      const raw = localStorage.getItem(key);
      return raw == null ? undefined : JSON.parse(raw);
    } catch (e) {
      return undefined;
    }
  }

  /* ----------------------------- loading ----------------------------- */

  async function loadJson(url) {
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error(`Failed to load ${url} (${res.status})`);
    return res.json();
  }

  async function loadAll() {
    setStatus("Loading rulesets…", true);
    try {
      index = await loadJson(INDEX_URL);

      // Restore the persisted selection before rendering the controls.
      const savedType = readSetting(SETTINGS_KEY_DOC_TYPE);
      const savedRulesets = readSetting(SETTINGS_KEY_RULESETS);
      if (host === "PowerPoint") {
        docType = "presentation";
      } else if (savedType && index.documentTypes[savedType]) {
        docType = savedType;
      }
      activeRulesetIds = Array.isArray(savedRulesets) && savedRulesets.length ? savedRulesets : null;

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
    // onchange (not addEventListener) so repeated loadAll() reloads never
    // accumulate duplicate handlers.
    sel.onchange = () => {
      docType = sel.value;
      activeRulesetIds = null;
      saveSetting(SETTINGS_KEY_DOC_TYPE, docType);
      saveSetting(SETTINGS_KEY_RULESETS, []);
      renderRulesetChips();
    };
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
      chip.className = "chip" + (!activeRulesetIds || activeRulesetIds.includes(rs.id) ? " active" : "");
      chip.textContent = rs.title.split("—")[0].split("(")[0].trim() || rs.id;
      chip.dataset.id = rs.id;
      chip.addEventListener("click", () => {
        chip.classList.toggle("active");
        const active = wrap.querySelectorAll(".chip.active");
        activeRulesetIds = active.length === available.length ? null : Array.from(active).map((c) => c.dataset.id);
        saveSetting(SETTINGS_KEY_RULESETS, activeRulesetIds || []);
      });
      wrap.appendChild(chip);
    }
  }

  function selectedRulesets() {
    let list = rulesetsForDocType();
    if (activeRulesetIds) list = list.filter((r) => activeRulesetIds.includes(r.id));
    return list;
  }

  /* ----------------------------- rendering ----------------------------- */

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
    const locText = (l) => (l.snippet != null && l.snippet !== "" ? l.snippet : `paragraph ${l.paragraphIndex}`);
    const locs = f.locations || [];
    const firstLoc = locs.length ? `<div class="finding-loc">${esc(locText(locs[0]))}</div>` : "";
    const moreLocs = locs.length > 1
      ? `<details class="finding-more"><summary>${locs.length - 1} more location(s)</summary>${locs.slice(1).map((l) => `<div class="finding-loc">${esc(locText(l))}</div>`).join("")}</details>`
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
      ${firstLoc}
      ${moreLocs}
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

    const fixable = report.findings.filter((f) => f.fix && f.fix.action).length;
    const fa = $("fix-all");
    if (fixable) {
      fa.hidden = false;
      fa.textContent = `Fix all (${fixable})`;
    } else {
      fa.hidden = true;
    }
  }

  /* ----------------------------- report export ----------------------------- */

  function buildReportText(report) {
    const typeLabel = index && index.documentTypes[report.docType]
      ? index.documentTypes[report.docType].label
      : report.docType;
    const c = report.counts;
    const lines = [];
    lines.push("DFSWM COMPLIANCE REPORT");
    lines.push("=".repeat(28));
    lines.push(`Document type : ${typeLabel}`);
    lines.push(`Generated     : ${new Date().toLocaleString()}`);
    lines.push(`Summary       : ${c.error} error(s), ${c.warning} warning(s), ${c.info} info, ${c.manual} manual`);
    lines.push("");
    const { groups, order } = groupBySeverity(report.findings);
    const labels = { error: "ERRORS", warning: "WARNINGS", info: "INFORMATION", manual: "MANUAL REVIEW" };
    for (const key of order) {
      const list = groups[key];
      if (!list.length) continue;
      lines.push(`${labels[key]} (${list.length})`);
      lines.push("-".repeat(28));
      list.forEach((f, i) => {
        lines.push(`${i + 1}. ${f.title}`);
        if (f.chapter) lines.push(`   Reference: Ch ${f.chapter}${f.paragraphs ? ` \u00b6${f.paragraphs}` : ""}`);
        lines.push(`   ${f.message}`);
        for (const l of f.locations || []) {
          lines.push(`   - ${l.snippet != null && l.snippet !== "" ? l.snippet : `paragraph ${l.paragraphIndex}`}`);
        }
      });
      lines.push("");
    }
    return lines.join("\n");
  }

  function exportReport() {
    if (!lastReport) return;
    const text = buildReportText(lastReport);
    try {
      const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `dfswm-report-${new Date().toISOString().slice(0, 10)}.txt`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 4000);
      setStatus("Report exported.");
    } catch (err) {
      // Some Office webviews block downloads; copy to the clipboard instead.
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text)
          .then(() => setStatus("Report copied to clipboard."))
          .catch(() => setStatus("Export failed."));
      } else {
        setStatus(`Export failed: ${err.message}`);
      }
    }
  }

  /* ----------------------------- progress ----------------------------- */

  function showBusy(on) {
    const wrap = $("progress-wrap");
    if (!wrap) return;
    wrap.hidden = !on;
    wrap.classList.toggle("busy", on);
    if (!on) $("progress-bar").style.width = "";
  }

  function setProgress(frac) {
    const bar = $("progress-bar");
    if (!bar) return;
    bar.style.width = `${Math.max(0, Math.min(1, frac)) * 100}%`;
  }

  /* ----------------------------- actions ----------------------------- */

  async function runCheck() {
    setStatus("Checking document…", true);
    $("run-check").disabled = true;
    showBusy(true);
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
      showBusy(false);
    }
  }

  async function onFixClick(e) {
    const btn = e.target.closest(".btn-fix");
    if (!btn || !lastReport) return;
    const idx = Number(btn.dataset.fixIndex);
    const findingObj = lastReport.findings.find((f) => f._fixIndex === idx);
    if (!findingObj || !findingObj.fix) return;
    setStatus("Applying fix…", true);
    showBusy(true);
    try {
      const result = await applyFix(findingObj.fix);
      setStatus(result.ok ? `Fix applied: ${result.message}` : `Fix failed: ${result.message}`);
      if (result.ok) setTimeout(runCheck, 300);
    } finally {
      showBusy(false);
    }
  }

  async function onFixAllClick() {
    const fixable = lastReport.findings.filter((f) => f.fix && f.fix.action);
    if (!fixable.length) return;
    setStatus("Applying fixes…", true);
    $("fix-all").disabled = true;
    showBusy(true);
    let okCount = 0;
    let failCount = 0;
    try {
      for (let i = 0; i < fixable.length; i++) {
        setProgress((i + 1) / fixable.length);
        const result = await applyFix(fixable[i].fix);
        if (result.ok) okCount += 1;
        else failCount += 1;
      }
      setStatus(`${okCount} fix(es) applied, ${failCount} failed`);
      if (okCount) setTimeout(runCheck, 300);
    } finally {
      $("fix-all").disabled = false;
      showBusy(false);
    }
  }

  function bind() {
    $("run-check").addEventListener("click", runCheck);
    $("back").addEventListener("click", () => {
      $("results-pane").hidden = true;
      $("config-pane").hidden = false;
    });
    $("load-rulesets").addEventListener("click", loadAll);
    $("results").addEventListener("click", onFixClick);
    $("fix-all").addEventListener("click", onFixAllClick);
    $("export").addEventListener("click", exportReport);
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