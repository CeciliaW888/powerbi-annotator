# Power BI Annotator

Chrome extension for annotating Power BI reports with drawings, comments, and professional exports.

![Version](https://img.shields.io/badge/version-1.0.0-blue) ![Chrome](https://img.shields.io/badge/chrome-extension-green)

---

## Features

- **5 Drawing Tools** - Rectangle, Arrow, Circle, Line, Freehand
- **Color Picker** - Choose any color for annotations
- **Sidebar Comments** - All annotations in one organized view with auto-numbering
- **Export to PDF** - HTML with embedded screenshot, print to PDF
- **Export to PowerPoint** - Direct .pptx download (opens in PowerPoint/Google Slides)
- **Export CSV** - Excel/CSV spreadsheet with matching numbers (#1 in PDF/PPT = Row 1 in CSV)
- **Auto-Save** - Annotations persist across sessions
- **Viewport Warning** - Alerts if annotations are off-screen before export

---

## Installation

1. Open Chrome and go to `chrome://extensions`
2. Enable **Developer mode** (toggle in top-right)
3. Click **Load unpacked**
4. Select the `powerbi-annotator` folder
5. Done! Open any Power BI report to use it

### Test It

1. Open `test-page.html` in Chrome (enable "Allow access to file URLs" in extension details)
2. Look for the **💬 button** on the right side
3. Click it, then click **Start Annotating** and try drawing

---

## How to Use

### Annotating

1. Go to **app.powerbi.com** and open a report
2. Click the **💬 button** on the right side
3. Click **Start Annotating** and select a tool + color
4. Click and drag to draw, then add your comment (Ctrl+Enter to submit)

### Managing Comments

- **Highlight** - Jump to any annotation on the page
- **Delete** - Remove individual annotations (badges renumber automatically)
- **Clear** - Delete all annotations at once

### Exporting

**Export Pages → PDF:** Click **Export Pages** → choose PDF → HTML downloads with embedded screenshot → open and print to PDF

**Export Pages → PowerPoint:** Click **Export Pages** → choose PowerPoint → `.pptx` file downloads directly → open in PowerPoint or Google Slides

**Export CSV (Excel):** Click **Export CSV** → CSV downloads automatically → open in Excel

---

## Project Structure

```
powerbi-annotator/
├── manifest.json              # Extension config
├── src/
│   ├── content/
│   │   ├── content.js        # Main logic
│   │   └── content.css       # Styles
│   ├── background/
│   │   └── background.js     # Background worker
│   └── lib/
│       └── pptxgen.bundle.js # PptxGenJS library for .pptx export
└── assets/icons/             # Extension icons
```

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Extension doesn't appear | Refresh page (F5), check it's enabled at `chrome://extensions` |
| Can't see 💬 button | Must be on app.powerbi.com, try scrolling, check for conflicting extensions |
| Annotations not saving | Check Chrome storage permissions, reinstall extension |
| Screenshot capture failed | Reload extension at `chrome://extensions` (click 🔄), refresh the page |
| Drawing toolbar hidden | Click "Start Annotating" first |

---

## Browser Support

- Chrome (recommended), Edge (Chromium), Brave
- Not supported: Firefox, Safari

## Privacy

All data stored locally in Chrome. No external servers, no tracking.

---

## Version History

**1.0.0** - Initial release with 5 drawing tools, direct PowerPoint export, PDF export, CSV export, activeTab screenshot flow, and organized structure

---

**Built for better Power BI collaboration**
