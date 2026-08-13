#!/usr/bin/env node
/**
 * Zero-dependency static server for local add-in development.
 * Serves dfswm-plugin/ (task pane, assets, synced rulesets) on
 * http://localhost:3000 — the URL used in manifest.xml.
 *
 * NOTE: Office sideloading normally requires HTTPS. For pure local UI/engine
 * testing, plain HTTP on localhost works in a browser. To sideload into Word,
 * either use `npx office-addin-dev-certs` + HTTPS, or follow the README.
 */
"use strict";

const http = require("http");
const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const PORT = Number(process.env.PORT || 3000);

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".xml": "application/xml; charset=utf-8",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
};

const server = http.createServer((req, res) => {
  let urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
  if (urlPath === "/") urlPath = "/src/taskpane.html";

  const filePath = path.normalize(path.join(ROOT, urlPath));
  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain" });
      res.end(`Not found: ${urlPath}`);
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, {
      "Content-Type": MIME[ext] || "application/octet-stream",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*",
    });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log(`DFSWM add-in dev server: http://localhost:${PORT}`);
  console.log(`  Task pane : http://localhost:${PORT}/src/taskpane.html`);
  console.log(`  Manifest  : dfswm-plugin/manifest.xml (edit URLs if you switch to HTTPS)`);
});
