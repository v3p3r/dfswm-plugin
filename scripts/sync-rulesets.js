#!/usr/bin/env node
/**
 * Copies the DFSWM rulesets (from ../dfswm-rulesets) into the plugin's
 * static web root so the task pane can fetch them via HTTP.
 *
 * Layout: dfswm-plugin/rulesets/<ruleset>.json and dfswm-plugin/index.json
 */
"use strict";

const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const SRC = path.resolve(ROOT, "..", "dfswm-rulesets");
const DEST = path.join(ROOT, "rulesets");

function main() {
  if (!fs.existsSync(SRC)) {
    console.error(`Rulesets source not found: ${SRC}`);
    console.error("Run this script from the project root that contains both dfswm-rulesets/ and dfswm-plugin/.");
    process.exit(1);
  }
  fs.mkdirSync(DEST, { recursive: true });

  const rulesetsDir = path.join(SRC, "rulesets");
  let copied = 0;
  for (const f of fs.readdirSync(rulesetsDir)) {
    if (!f.endsWith(".json")) continue;
    fs.copyFileSync(path.join(rulesetsDir, f), path.join(DEST, f));
    copied += 1;
  }
  fs.copyFileSync(path.join(SRC, "index.json"), path.join(ROOT, "index.json"));

  console.log(`Synced ${copied} rulesets + index.json into ${path.relative(ROOT, DEST)}/`);
}

main();
