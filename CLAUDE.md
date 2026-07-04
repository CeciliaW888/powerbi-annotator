# CLAUDE.md

Guidance for Claude Code (and other AI agents) working in this repository.

## Project Overview

**Power BI Annotator** — Chrome Manifest V3 extension that adds annotation capabilities to Power BI Service reports. Vanilla JS, no build step, no framework. Drawing tools, per-page comment storage, PDF/PPT/Excel export.

## Workflow Rules (READ FIRST)

1. **Wait for user confirmation before editing.** When you present a plan, options, or trade-offs, stop and wait. Do not start editing until the user says to proceed.
2. **No build step.** Edit files, reload extension at `chrome://extensions` (🔄), F5 the Power BI tab. "Extension context invalidated" errors mean the page wasn't refreshed after reload.
3. **Tests live in `tests/` and run with `npm test`** (`node --test`). They use Node's built-in test runner plus `jsdom`. 35 tests currently pass.
4. **Before claiming a fix works, run the tests AND verify in a browser** if the change touches drawing, navigation, storage, or export. The Playwright skill has been used previously — see "Browser testing" below.
5. **Norton/corporate SSL interception** can break `npm install`/`git clone`. Workarounds the user has already applied: `NODE_OPTIONS=--use-system-ca` for node, `git config --global http.sslBackend schannel` for git.

## Architecture

### Module Map (manifest.json content_scripts load order)

The order matters — earlier scripts must be loaded before later ones reference them.

```
1. src/lib/pptxgen.bundle.js          PptxGenJS (.pptx)
2. src/lib/jspdf.umd.min.js           jsPDF (.pdf)
3. src/lib/xlsx.full.min.js           SheetJS (.xlsx)
4. src/content/coords.js              Pure canvas-relative coordinate conversion + v1 migration
5. src/content/tools.js               Pure drawing-tool rendering + geometry
6. src/content/page-store.js          Per-page annotation storage (canonical keys + SPA nav)
7. src/content/page-navigator.js      Pure nav-element finder (workspace tabs + App left-nav)
8. src/content/presentation-layout.js Pure export layout math
9. src/content/content.js             Main orchestration — UI, drawing handlers, exports
```

Each of the three new modules attaches to `window.PowerBIAnnotator*` for the browser and also `module.exports`s for the Node tests via `_setup.js` (jsdom DOM globals).

### Page-world script

`src/content/powerbi-page-script.js` is **not** a content script — it's injected via `web_accessible_resources` into the **page world**, where it can:
- Access `powerbi.embeds[0].getActivePage()` (the Power BI Embed API isn't available in the content-script isolated world)
- Intercept `history.pushState` / `replaceState` (which content scripts can't override across the isolation boundary)

It communicates back to `content.js` via `window.postMessage` with two message types:
- `__pbi_annotator_page_info__` — current Power BI page name + URL
- `__pbi_annotator_navigation__` — fires the instant pushState fires (used for fast SPA nav detection)

### Module: tools.js

Exposes `window.PowerBIAnnotatorTools = { rectangle, line, arrow, circle, freehand, computeGeometry, geometryFromAnnotation }`.

- Each tool: `{ name, render(geometry, color) }` returning an SVG element (or `null` for rectangle, whose box border IS the rectangle)
- `computeGeometry(toolName, startPoint, currentPoint, freehandPoints?)` — pure function, returns `{ x, y, width, height, x1, y1, x2, y2, freehandPath? }`
- `geometryFromAnnotation(annotation)` — normalizes freehand path coords from **absolute** (page-origin) to **box-local** using `Math.min`. This was caught via a RED test during the refactor; do not "fix" it by assuming box-local.

### Module: page-store.js

Factory: `createPageStore({ storage, locationProvider, displayNameResolver, pageOrderResolver })` — pure dependency injection so tests can swap real `chrome.storage` for an in-memory map.

**Hybrid page key:** `deriveKeyFromLocation(loc)` returns
- `${reportId}#${sectionHash}` when a Power BI URL exposes both — survives SPA navigation between pages
- Else falls back to `pathname + search` — covers test-page.html and legacy data

`init()` reads `chrome.storage.local`, migrates legacy pathname-only keys to canonical keys, and populates an internal `cache` object.

**The cache is a `const`** — mutated in place via delete-then-assign so external mirror references (`allAnnotationsCache` in `content.js`) stay valid. Don't reassign it.

`_snapshot()` returns the live mirror reference; `content.js` uses it to keep its legacy `allAnnotationsCache` working.

Methods: `init`, `current`, `list`, `saveAnnotations`, `deleteAnnotations`, `deleteAll`, `onPageChange`, `onDataChange`, `checkPageChange`.

### Module: presentation-layout.js

Pure layout helpers, no DOM:
- `SLIDE_DIMENSIONS` — `{ width: 13.33, height: 7.5 }` (PPT widescreen)
- `computeImageFit(srcW, srcH, boxW, boxH)` — contain-fit math; returns `{ width, height, offsetX, offsetY }`
- `commentsPerSlide({ availableHeight, lineHeight })` — integer
- `chunkComments(comments, maxPerSlide)` — splits comments across slides

`content.js` calls these in **four** export functions (PDF, PPT, two more). Don't duplicate the math inline — it was deduplicated in the refactor for a reason.

### content.js — what's still there

~2600 lines. Owns the UI (sidebar DOM creation), mouse-event drawing handlers, custom modal dialogs (`showModal` / `showConfirm` / `showPrompt`), export orchestration, comment-list rendering. Delegates to the three modules above.

Key state variables:
- `annotations[]` — annotations for the current page
- `allAnnotationsCache` — live mirror of `pageStore._snapshot()` (mirror reference; never reassign)
- `screenshotCache` — `{ [pageKey]: dataUrl }` of cached PNG screenshots
- `isAnnotationMode`, `currentDrawingTool`, `currentColor`, `sidebarOpen`

`getPageKey()` delegates to `pageStore.current().key`. **Do not** reintroduce a parallel implementation — the page-key-mismatch bug fixed in commit `09d3f40` came from exactly that.

### Annotation object shape

```js
{
  id: timestamp * 100 + counter,    // collision-free
  x, y, width, height,              // page-origin pixels
  comment: "user text",
  timestamp: ISO string,
  tool: "rectangle|arrow|circle|line|freehand",
  color: "#hex",
  freehandPath: [{x, y}],           // freehand only — ABSOLUTE coords
  startPoint: {x, y},               // arrow/line direction
  endPoint:   {x, y},

  // v2 (canvas-anchored) fields — added by coords.js, see recurring bug #7:
  coordSpace: 'canvas',             // present ⇒ render via resolveAnnotationForLayout
  rel:         {x, y, w, h},        // box as fractions (0–1) of the report canvas rect
  relStart:    {x, y},              // relStart/relEnd/relFreehand mirror the pixel points
  relEnd:      {x, y},
  relFreehand: [{x, y}] | null
}
```

Legacy v1 annotations (no `coordSpace`) are migrated to v2 on first render where the canvas is found (`migrateLoadedAnnotations`), using the current canvas rect. The absolute `x/y/width/height` are always kept as a v1 fallback.

## Recurring Bugs — Verify After Every Refactor

These bugs have been re-introduced multiple times during merges/rollbacks. Before claiming work is done, **verify each one is still fixed** (grep + run tests + run Playwright if any module changes):

1. **postMessage navigation listener** — `content.js` must listen for `__pbi_annotator_navigation__` and call `onPageChanged()`. Without it, SPA nav only fires on the 1s polling fallback, and the user sees drawings on the wrong page.
2. **Global numbering by KEY only** — `getGlobalStartNumber` must match by `page.key === currentKey`. Do **NOT** also match `page.name === currentPageName` — test pages share page names, breaking the loop.
3. **Report ID format match** — `getAnnotatedPages` must normalize via `storedPath.split('/').filter(p => p).join('_')` to match `getReportId()`'s format. Raw `storedPath` mismatches and filters out other pages.
4. **Screenshot caching timing** — Do NOT call `cacheCurrentScreenshot(oldKey)` in `onPageChanged` (annotations are already cleared from DOM). DO cache after rendering on the new page (after the 400ms DOM-settle timeout). Hide sidebar AND toggle button before capture, with a 200ms repaint delay.
5. **`captureVisibleTab` requires `activeTab` or `<all_urls>`** — `chrome.tabs.captureVisibleTab` from the background worker fails on file:// AND https:// when only specific `host_permissions` are set. Error: `"Either the '<all_urls>' or 'activeTab' permission is required."` This is a Chrome limitation; silent caching is impossible without `<all_urls>` (which triggers stricter Web Store review). Workaround: the user clicks the extension icon to grant `activeTab` per-export.
6. **PageStore key vs. `getPageKey()` mismatch** — `getPageKey()` MUST delegate to `pageStore.current().key`. If it returns `pathname+search` while PageStore writes canonical `reportId#sectionHash`, you'll get two coexisting key formats and annotations disappear (commit `09d3f40` fix).
7. **Never read `annotation.x/y` directly for display** — annotations must always be rendered through `resolveAnnotationForLayout` (which rescales v2 `rel` fractions to the current canvas rect). Reading the stored absolute `x/y` bypasses the canvas anchoring and reintroduces the drift-on-layout-change bug. The stored `x/y/width/height` are kept only as a v1 fallback / rollback safety net.

## Screenshot / Export Flow

### Single page (`generatePresentation`)
1. Hides sidebar + toggle button, sends `{ action: 'prepareCapture' }` to background.
2. `waitForScreenshotOrCancel()` shows a modal: "Click the extension icon".
3. User clicks the extension icon → `activeTab` grants → background `chrome.tabs.captureVisibleTab`.
4. Background returns `{ action: 'screenshotResult', screenshot: dataUrl }`; content bundles into PDF/PPT.

### Multi-page (`generateMultiPagePresentation` — guided wizard)
The old per-page `screenshotCache` reliance is retired. The wizard walks every annotated page live:
1. `getAnnotatedPages()` yields the pages; a progress modal (`showExportProgress`) lists them.
2. For each page not already current, `PageNavigator.findNavElement` locates the tab/nav element and clicks it, then `waitForPageSettle` waits for the key + canvas to settle.
3. `captureVisiblePage()` tries a silent `captureForCache` first; on the first permission failure it falls back to the icon-click modal ONCE. `activeTab`, once granted, survives Power BI's `pushState` page switches, so every later capture in the loop is silent.
4. Each capture is cropped via `cropScreenshotToCanvas`, assembled into `pageDataList` (`{ pageName, screenshot, comments }`, comments globally numbered), and handed to `generateMultiPagePdf` / `generateMultiPagePptx`.
5. Pages whose nav element can't be found are marked **failed** in the modal rather than exported blank.

Excel export (`buildExcelData` + SheetJS) needs no screenshots. `cacheCurrentScreenshot` still exists for the single-page nicety but the wizard does not depend on it.

## Common Tasks

### Add a new drawing tool

1. Add button to toolbar HTML in `createSidebar()`
2. Add an entry to `window.PowerBIAnnotatorTools` in `tools.js` with `{ name, render(geometry, color) }`
3. Extend `computeGeometry` if the tool needs special geometry
4. Add tests in `tests/tools.test.js`
5. `createAnnotationElement` and `updateAnnotationVisual` will pick it up automatically via the Tools delegation

### Change export layout

- **PDF**: HTML template inside `generatePresentation()` (the `format === 'pdf'` path)
- **PPT**: `generatePptx()` — uses PptxGenJS, image fit via `computeImageFit`
- **Excel**: `buildExcelData()` (headers + rows) and the `ws['!cols']` width array

Don't inline image-fit math — use `window.PowerBIAnnotatorPresentationLayout.computeImageFit`.

### Change storage behavior

- All reads/writes go through `pageStore` — do not call `chrome.storage.local` directly from `content.js`
- Migration of legacy pathname-only keys runs in `pageStore.init()`
- `allAnnotationsCache` is a live mirror; never reassign it

## Testing

```bash
npm test                                  # 35 tests, ~1s
```

Tests live in `tests/`:
- `tools.test.js` — tool rendering + geometry
- `page-store.test.js` — storage, key derivation, migration
- `presentation-layout.test.js` — image fit + comment chunking
- `_setup.js` — jsdom bootstrap (DOM globals)

### Browser testing

The Playwright skill at `C:\Users\cecil\.claude\skills\playwright-skill\run.js` can launch Chrome with the extension loaded. Two test scripts exist:

- `C:\Users\cecil\AppData\Local\Temp\playwright-test-pbi-extension.js` — basic carry-over check
- `C:\Users\cecil\AppData\Local\Temp\playwright-test-recurring-bugs.js` — 9-scenario test covering all five tools, badge numbering, page switching, sidebar aggregation, instant nav, reload persistence

Notes for writing new Playwright tests:
- Tool buttons can render outside the visible viewport — use `page.evaluate(() => document.querySelector('[data-tool="..."]').click())` instead of `.click()`
- Test page selectors: `#pbi-toggle-btn` (toggle), `.pbi-annotation-box`, `.pbi-annotation-number`, `.pbi-comment-item`, `[data-tool="<name>"]`
- The test page (`test-page.html`) uses plain `<button>` clicks for tab nav — `history.pushState` is NOT called, so the `__pbi_annotator_navigation__` postMessage path won't fire there. The 1s polling fallback will fire instead. This is expected on the test page; against real Power BI, postMessage fires instantly.

Run:
```bash
cd "C:\Users\cecil\.claude\skills\playwright-skill"
NODE_OPTIONS="--use-system-ca" node run.js "<absolute path to test script>"
```

## Permissions Explained

```json
"permissions": ["storage", "activeTab", "tabs"]
```
- `storage` — annotations and screenshot cache
- `activeTab` — temporary screenshot permission on extension-icon click
- `tabs` — required for `chrome.tabs.captureVisibleTab`

```json
"host_permissions": ["https://app.powerbi.com/*", "https://*.powerbi.com/*", "file:///*"]
```
`file:///*` is for local testing with `test-page.html`. Do not add `<all_urls>` without weighing the Web Store review impact.

## Power BI Specifics

- URLs change per page: `.../ReportSection{sectionHash}` — handled by PageStore's canonical key
- Visuals render to HTML canvas; annotations are absolute pageX/pageY overlays, not DOM-anchored — layout shifts can misalign them
- Republishing a report can change section IDs — annotations attached to the old IDs orphan (Power BI limitation)
- The Embed API is page-world only — that's why `powerbi-page-script.js` exists

## What NOT to Do

- Don't inline drawing SVG creation in `content.js` — use `Tools`
- Don't inline page-key derivation — use `pageStore.current().key`
- Don't reassign `allAnnotationsCache` or the PageStore `cache` — both are live mirrors
- Don't add `<all_urls>` without discussing the Web Store review trade-off with the user first
- Don't add an automatic re-capture loop that polls `captureVisibleTab` from the background worker — it requires `activeTab` (user gesture) on all URLs
- Don't add error handling for impossible scenarios; validate only at boundaries (user input, chrome APIs)
- Don't write comments explaining WHAT — only WHY when non-obvious
