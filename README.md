# Power BI Annotator

Chrome extension for annotating Power BI reports. Draw on dashboards, leave comments, and export to PDF, PowerPoint, or Excel.

![Version](https://img.shields.io/badge/version-1.2.0-blue) ![Chrome](https://img.shields.io/badge/chrome-extension-green)

---

## Features

- **5 Drawing Tools** — Rectangle, Arrow, Circle, Line, Freehand
- **Color Picker** — Pick any color per annotation
- **Smart Numbering** — Annotations numbered globally across all pages (#1, #2, #3…)
- **Multi-Page Support** — Switch between report pages; annotations stay with their page
- **Multi-Report Support** — Each report is scoped separately
- **Continuous Annotation Mode** — Stays on as you navigate
- **Sidebar** — All comments in one view, grouped by page
- **Export to PDF** — Annotated dashboard + numbered comment list
- **Export to PowerPoint** — Real `.pptx` file with widescreen slides
- **Export to Excel** — Spreadsheet with page URLs and comments
- **Auto-Save** — Everything is saved locally in Chrome storage
- **Off-Screen Warning** — Alerts you before exporting if annotations would be cropped

---

## Installation

1. Open Chrome and go to `chrome://extensions`
2. Enable **Developer mode** (toggle in top-right)
3. Click **Load unpacked**
4. Select the `powerbi-annotator` folder
5. Done — open any Power BI report to use it

After updating the code:
1. Click the reload icon (🔄) on the extension card at `chrome://extensions`
2. **Press F5** on any open Power BI tab (otherwise you'll get an "Extension context invalidated" error)

---

## How to Use

### Annotating

1. Open a report at **app.powerbi.com**
2. Click the **💬 button** on the right edge
3. Click **Start Annotating**, then pick a tool and color
4. Click-and-drag on the page to draw, then type your comment (Ctrl+Enter to submit)
5. Sidebar auto-hides during drawing — click 💬 to reopen it
6. Navigate between report pages freely — annotation mode stays on
7. Annotations are numbered globally (Page 1: #1–3, Page 2: #4–6, etc.)

### Managing Comments

- **Highlight** — Jumps to and flashes the annotation on the page
- **Delete** — Removes one annotation (remaining badges renumber)
- **Clear Page** — Removes all annotations on the current page
- **Clear All** — Removes annotations across every page of the report

### Exporting

**Export PDF / PPT**
1. Click **Export PDF / PPT** → choose **PDF** or **PowerPoint**
2. If you chose **All pages**, a progress panel opens and the extension walks each annotated page automatically, capturing it as it goes
3. The **first** capture asks you to click the **💬 extension icon** in the Chrome toolbar (top-right) to grant screenshot permission; every page after that is captured silently
4. The file downloads automatically when the walk finishes (open `.pptx` in PowerPoint, Google Slides, or LibreOffice Impress)

**Export Excel**
1. Click **Export Excel** — file downloads immediately (no icon click needed)
2. Columns: No, Page Name, URL, Date, Comment
3. Numbers match the PDF/PPT exports

---

## Multi-Page Export Workflow

Chrome's screenshot API can only capture the page you're currently looking at. The multi-page export is now a **guided wizard** that handles this for you: it navigates to each annotated page in turn, waits for it to render, and captures it — no more manual page-by-page pre-caching.

### What happens when you export all pages

1. Draw annotations on **every page** you want in the export
2. Click **Export PDF / PPT** → **PDF** or **PowerPoint** → **All pages**
3. A progress panel lists each page and marks it **active → done** as it is captured
4. The **first** page prompts you to click the **💬 extension icon** to grant screenshot permission; the wizard then captures the remaining pages silently
5. When every page is done, the file downloads and you are returned to the page you started on

If the wizard can't find a page's navigation tab, that page is marked **failed** in the panel and skipped (rather than exported blank). You can cancel mid-run with the **Cancel** button.

---

## Other Limitations

- **Power BI report republishing** — If a report is republished and its internal page IDs change, annotations attached to the old IDs will not appear (this is a Power BI platform limitation, not a bug in the extension)
- **Layout shifts** — Annotations are anchored to the report canvas as relative fractions, so they follow the canvas across window resizes, sidebar toggles, and App view. Extreme zoom or a republished layout can still misalign them
- **Export captures only the visible viewport** — Scroll to include off-screen content, or use the "off-screen annotation" warning before exporting
- **Canvas-rendered visuals** — Annotations sit on top of the page as an overlay; they don't bind to specific Power BI visuals
- **Chrome / Edge only** — Firefox uses a different extension API and is not supported

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Extension doesn't appear | Refresh page (F5); confirm it's enabled at `chrome://extensions` |
| Can't see 💬 button | Must be on app.powerbi.com; try scrolling; check for conflicting extensions |
| Annotations not saving | Check Chrome storage permissions, reinstall extension |
| Drawings showing on the wrong page | Refresh the page (F5). If it persists, file a bug — this is a known recurring issue |
| "Extension context invalidated" error | You reloaded the extension — press **F5** on the Power BI tab |
| "Message port closed" error | Extension communication issue — refresh the page and try again |
| Screenshot capture failed | Reload extension at `chrome://extensions`, then F5 the page |
| Export stuck waiting for screenshot | Click the **extension icon** (💬 in Chrome toolbar), not a page button |
| Exported PDF has a page marked "failed" | The wizard couldn't find that page's navigation tab. Make sure the report's page tabs are visible, then re-export. See *Multi-Page Export Workflow* above |
| Drawing toolbar hidden | Click "Start annotating" — the sidebar auto-hides to give full screen space |
| Annotations from a different report appear | Each report is scoped separately and shouldn't bleed. Refresh and report a bug if it happens |

---

## Privacy

All data — annotations and cached screenshots — is stored locally in Chrome's `storage.local`. Nothing is sent to any server. The extension never makes outbound network requests.

---

## Project Structure

```
powerbi-annotator/
├── manifest.json                    # Extension config (entry point)
├── src/
│   ├── background/
│   │   └── background.js            # Screenshot capture service worker
│   ├── content/
│   │   ├── content.js               # Main UI + drawing + export orchestration
│   │   ├── content.css              # Sidebar & annotation styles
│   │   ├── coords.js                # Canvas-relative coordinate conversion + v1 migration (pure module)
│   │   ├── tools.js                 # Drawing tool rendering + geometry (pure module)
│   │   ├── page-store.js            # Per-page annotation storage with SPA-nav awareness
│   │   ├── page-navigator.js        # Finds page nav elements for the export wizard (pure module)
│   │   ├── presentation-layout.js   # Pure helpers for export image fit + comment paging
│   │   └── powerbi-page-script.js   # Injected into page world for Power BI Embed API + pushState hook
│   └── lib/
│       ├── pptxgen.bundle.js        # PptxGenJS — .pptx export
│       ├── jspdf.umd.min.js         # jsPDF — .pdf export
│       └── xlsx.full.min.js         # SheetJS — .xlsx export
├── tests/                           # Node test runner + jsdom (43 tests)
├── package.json                     # Dev dependency: jsdom
└── README.md
```

No build step — edit and reload.

---

## Browser Support

- ✅ Chrome (primary target)
- ✅ Edge (Chromium)
- ✅ Brave
- ❌ Firefox (different extension API)
- ❌ Safari

Requires Manifest V3.

---

## Version History

**v1.3** — Canvas-anchored annotations (survive layout changes + App view), guided multi-page export wizard, Fluent-style UI refresh (color swatches, toasts, progress panel)
**v1.2** — Excel export with page URLs, improved page detection, multi-page support, refactored into testable modules
**v1.1** — Smart numbering across pages, continuous annotation mode
**v1.0** — Initial release

---

**Built for better Power BI collaboration**
