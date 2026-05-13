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

**Export Pages → PDF**
1. Click **Export Pages** → choose **PDF**
2. A modal asks you to click the extension icon
3. Click the **💬 extension icon** in the Chrome toolbar (top-right)
4. A `.pdf` file downloads automatically

**Export Pages → PowerPoint**
1. Click **Export Pages** → choose **PowerPoint**
2. Click the **💬 extension icon** when prompted
3. A `.pptx` file downloads — open in PowerPoint, Google Slides, or LibreOffice Impress

**Export Excel**
1. Click **Export Excel** — file downloads immediately (no icon click needed)
2. Columns: No, Page Name, URL, Date, Comment
3. Numbers match the PDF/PPT exports

---

## ⚠️ Important: Multi-Page Export Workflow

Chrome's screenshot API can only capture the page you're currently looking at. To make multi-page exports work, the extension **caches a screenshot of each page** after you draw on it and navigate away. A **📸 camera icon** appears in the sidebar next to a page once its screenshot is cached.

**If you export before all pages have a 📸, those pages will be blank in the output.**

### Recommended workflow

1. Draw annotations on **every page** you want in the export
2. Wait until a **📸** appears next to every page in the sidebar
   - Caching triggers after you navigate **away** from a page that has annotations rendered, so visit each page, then navigate to the next
3. *Then* click **Export Pages → PDF** or **PowerPoint**
4. When prompted, click the extension icon — this captures the page you're currently on (the only one missing a 📸)

If a page is missing a 📸 at export time, that slide will be blank or say "No screenshot available."

---

## Other Limitations

- **Power BI report republishing** — If a report is republished and its internal page IDs change, annotations attached to the old IDs will not appear (this is a Power BI platform limitation, not a bug in the extension)
- **Layout shifts** — Annotations use absolute pixel positions. Significant resizing, zooming, or theme changes after annotating may make them appear misaligned
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
| Exported PDF has blank pages | Some pages were missing the 📸 icon at export time. See *Multi-Page Export Workflow* above |
| Drawing toolbar hidden | Click "Start Annotating" — the sidebar auto-hides to give full screen space |
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
│   │   ├── tools.js                 # Drawing tool rendering + geometry (pure module)
│   │   ├── page-store.js            # Per-page annotation storage with SPA-nav awareness
│   │   ├── presentation-layout.js   # Pure helpers for export image fit + comment paging
│   │   └── powerbi-page-script.js   # Injected into page world for Power BI Embed API + pushState hook
│   └── lib/
│       ├── pptxgen.bundle.js        # PptxGenJS — .pptx export
│       ├── jspdf.umd.min.js         # jsPDF — .pdf export
│       └── xlsx.full.min.js         # SheetJS — .xlsx export
├── tests/                           # Node test runner + jsdom (35 tests)
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

**v1.2** — Excel export with page URLs, improved page detection, multi-page support, refactored into testable modules
**v1.1** — Smart numbering across pages, continuous annotation mode
**v1.0** — Initial release

---

**Built for better Power BI collaboration**
