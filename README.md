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

**Export Pages → PDF:**
1. Click **Export Pages** → choose **PDF**
2. A modal will appear asking you to click the extension icon
3. **Click the 💬 extension icon** in the Chrome toolbar (top-right)
4. Wait for the `.pdf` file to download directly
5. Open the PDF file

**Export Pages → PowerPoint:**
1. Click **Export Pages** → choose **PowerPoint**
2. A modal will appear asking you to click the extension icon
3. **Click the 💬 extension icon** in the Chrome toolbar (top-right)
4. Wait for the `.pptx` file to download directly
5. Open in PowerPoint or Google Slides

**Export CSV (Excel):**
1. Click **Export CSV**
2. CSV downloads automatically (no icon click needed)
3. Open in Excel - numbers match PDF/PPT annotations (#1, #2, etc.)

**Important Notes:**
- For PDF/PowerPoint export, you **must click the extension icon** when prompted (this grants screenshot permission)
- If you just reloaded/updated the extension, **refresh the page** before exporting
- Export captures only the visible viewport - scroll to include off-screen content

---

## Project Structure

```
powerbi-annotator/
├── assets/icons/              # Extension icons
├── src/
│   ├── background/
│   │   └── background.js      # Background worker
│   ├── content/
│   │   ├── content.css        # Styles
│   │   └── content.js         # Main logic
│   └── lib/
│       ├── jspdf.umd.min.js   # jsPDF library for .pdf export
│       └── pptxgen.bundle.js  # PptxGenJS library for .pptx export
├── manifest.json              # Extension config
└── README.md                  # Documentation
```

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Extension doesn't appear | Refresh page (F5), check it's enabled at `chrome://extensions` |
| Can't see 💬 button | Must be on app.powerbi.com, try scrolling, check for conflicting extensions |
| Annotations not saving | Check Chrome storage permissions, reinstall extension |
| "Extension context invalidated" error | You reloaded the extension - **refresh the page (F5)** before exporting |
| "Message port closed" error | Extension communication issue - **refresh the page** and try again |
| Screenshot capture failed | Reload extension at `chrome://extensions` (click 🔄), then refresh the page |
| Export stuck waiting for screenshot | Make sure you **click the extension icon** (💬) when the modal appears, not the page button |
| Drawing toolbar hidden | Click "Start Annotating" first |

---

## Browser Support

- Chrome (recommended), Edge (Chromium), Brave
- Not supported: Firefox, Safari

## Privacy

All data stored locally in Chrome. No external servers, no tracking.

---

## Version History

**1.0.0** - Initial release with 5 drawing tools, direct PowerPoint export, direct PDF export, CSV export, activeTab screenshot flow, and organized structure

---

**Built for better Power BI collaboration**
