# DFSWM Compliance Plugin

A Microsoft Office add-in (Word **and** PowerPoint) that checks the open
document against the **Kenya Defence Forces Service Writing Manual 2024
(JP-04)** rulesets and offers one-click autofixes.

The add-in loads the rulesets for the selected document type (letters,
memos, service papers, minutes, orders, briefs, presentations, …), extracts
the document into a plain, serialisable model via the Office JavaScript API,
runs a host-agnostic compliance engine over it, and renders the findings
grouped by severity — with targeted **Fix** buttons and a **Fix all** action.

## Features

- **Document type selection** — 20+ KDF document types, each mapped to the
  rulesets that apply (letters, memos, ESM, emails, signals, service papers,
  minutes/agenda, briefs, presentations, orders/FRAGO, …).
- **Ruleset toggles** — enable or disable individual rulesets per run.
- **Severity-grouped report** — errors, warnings, info and manual-review
  items, with chapter/paragraph references and snippet locations.
- **Targeted autofixes** — Fix buttons mutate **only the flagged
  paragraphs** (font, size, alignment, bold/italic, underline, line spacing,
  keep-together, heading formatting, regex text fixes), never the whole
  document. Document-level fixes (margins, page numbers, copy numbers) apply
  document-wide as designed.
- **Fix all** — apply every fixable finding in one pass with progress.
- **Report export** — download (or copy to clipboard) a plain-text report.
- **Settings persistence** — the selected document type and ruleset toggles
  are remembered in `Office.settings` (localStorage fallback) and restored
  when the task pane reopens.
- **Zero runtime dependencies** — everything (engine, bundler, icon
  generator, dev server) is plain Node/vanilla JS.

## Architecture

```
src/
  engine/                  host-agnostic core (runs in Node AND the browser)
    model.js               normalised document model + scope resolvers
    evaluators.js          one function per check type (regex, font-size, …)
    engine.js              runCompliance(): orchestrates rulesets + evaluators
  word/                    Word JavaScript API layer
    extractor.js           reads the open document into the model
    fixer.js               applies fix objects to Word
  powerpoint/              PowerPoint JavaScript API layer
    extractor.js           maps slides into the same model
    fixer.js               applies fix objects to slides
  taskpane.html/.css/.js   the task pane UI
  commands.html            ribbon command host page
rulesets/                  synced JSON rule definitions (see scripts/)
index.json                 document type → ruleset mapping (synced)
scripts/                   build tooling (zero dependencies)
  sync-rulesets.js         copy rulesets from ../dfswm-rulesets
  generate-icons.js        generate PNG add-in icons
  build-browser.js         bundle engine/extractors/fixers into dist/bundle.js
  dev-server.js            static server for local testing
test/                      node --test suites
```

### How a check runs

1. `taskpane.js` loads `index.json` + the ruleset JSON files.
2. The host extractor (`word/extractor.js` or `powerpoint/extractor.js`)
   reads the open document into a **plain model**: paragraphs (text, style,
   font, size, bold/italic/underline, alignment, indents, line spacing,
   keep-together, heading level, slide index), sections (margins,
   header/footer text), tables and footnotes.
3. `runCompliance(model, { docType, rulesets })` normalises the model once
   and runs every applicable rule through its evaluator.
4. Each finding carries `locations` (flagged paragraph indices + snippets)
   and, when autofixable, a `fix` object `{ action, params, locations }`.
   The engine copies the finding's locations onto the fix so fixers can
   target exactly the flagged paragraphs.
5. `taskpane.js` renders the findings; clicking **Fix** calls the host
   fixer with the fix object and re-runs the check.

## Prerequisites

- Node.js ≥ 18
- The `dfswm-rulesets` repository as a sibling folder
  (`../dfswm-rulesets`), which is the source of the ruleset JSON files.
- Microsoft Word / PowerPoint (for sideloading; desktop or web).
- Python 3 (only for `npm run validate:rulesets`, optional).

## Setup

No `npm install` is required — the project has no dependencies.

```bash
# 1. Sync the rulesets from ../dfswm-rulesets into rulesets/ + index.json
npm run sync:rulesets

# 2. Generate the ribbon icons
npm run icons

# 3. Bundle the engine/extractors/fixers for the browser
npm run build:browser

# All three at once:
npm run build

# Run the test suites (pure Node — no Office required)
npm test
```

### Local dev server

```bash
npm run dev            # builds, then serves http://localhost:3000
```

The manifest and the task pane reference `http://localhost:3000`. The dev
server also serves the task pane for plain-browser UI work:
http://localhost:3000/src/taskpane.html (the UI needs Office.js to run the
check, but the page loads and settings persist via localStorage).

### Sideloading into Word / PowerPoint

1. Run `npm run dev` (or serve the folder over HTTPS).
2. Office add-in sideloading normally requires **HTTPS**. Generate trusted
   certs with `npx office-addin-dev-certs install`, then serve the folder
   over the HTTPS port and update the URLs in `manifest.xml` (all
   `http://localhost:3000` occurrences, including `IconUrl`,
   `SourceLocation`, `AppDomain`, and the `bt:Image`/`bt:Url` resources).
3. Sideload the manifest:
   - **Word/PowerPoint desktop:** *File → Account → Manage Add-ins →
     Upload My Add-in* (or `npx office-addin-debugging start manifest.xml`).
   - **Word/PowerPoint on the web:** *Insert → Add-ins → Upload My Add-in*.
4. Open a document and click the **DFSWM → Check Compliance** ribbon button.

## Rulesets

Rulesets live in the separate `dfswm-rulesets` repository and are copied in
by `npm run sync:rulesets`. Each ruleset JSON has `id`, `title`, `chapter`
and a `rules[]` array. A rule looks like:

```json
{
  "id": "core.font",
  "title": "Font type must be Times New Roman or Arial",
  "severity": "error",
  "appliesTo": ["general"],
  "check": {
    "type": "font-name",
    "scope": "all",
    "params": { "allowed": ["Times New Roman", "Arial"] }
  },
  "autofix": { "available": true }
}
```

- `appliesTo` — `["general"]` for universal rules, otherwise the document
  type id(s) from `index.json`.
- `check.type` — must map to an evaluator in `src/engine/evaluators.js`
  (covered by a registry test). `check.scope` selects paragraphs via
  `paragraphsForScope()`: `body`, `all`, `subject-heading`, `main-heading`,
  `group-heading`, `paragraph-heading`, `heading`, `signature-block`,
  `first-page`, `header`, `footer`, `table`, `table-column-head`,
  `footnote`.
- `manual: true` renders the rule as a human checklist item.
- `autofix.available` enables the **Fix** button. Regex autofixes can
  declare `autofix.replacement`; known patterns (midnight → `0000 hours`,
  punctuation spacing, abbreviation dots) get sensible defaults in
  `regexFixFor()`.

### Adding a new rule

1. Add the rule to the ruleset JSON in `../dfswm-rulesets` and re-run
   `npm run sync:rulesets`.
2. If the check type is new, implement an evaluator in
   `src/engine/evaluators.js` (see the `EVALUATORS` registry) returning
   `Finding[]` via the `finding()` helper.
3. If the rule should be autofixable, emit a `fix` object with an `action`
   that one of the fixers implements (`src/word/fixer.js`,
   `src/powerpoint/fixer.js`), or add the action to both.
4. Add tests in `test/evaluators.test.js`, then `npm test` and
   `npm run build:browser`.

## Testing

```bash
npm test
```

Runs `node --test` over the engine, presentation-mapping and evaluator
suites. The suites use only the pure core — no Office runtime required. The
registry test also cross-checks that every `check.type` used by the synced
rulesets has an implemented evaluator.

## Troubleshooting

- **Task pane shows "Failed to load rulesets"** — run `npm run sync:rulesets`
  (the rulesets folder must exist alongside `dfswm-rulesets`).
- **Sideload rejects the manifest** — the add-in must be served over HTTPS;
  update every `http://localhost:3000` URL in `manifest.xml` to your HTTPS
  origin.
- **Office.settings not saving** — the task pane falls back to
  localStorage automatically.
- **A Fix button fails** — some actions are host-limited by design
  (e.g. page-number insertion in Word, margins in PowerPoint) and return an
  instructional message instead.