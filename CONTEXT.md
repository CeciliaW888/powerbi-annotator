# Power BI Annotator

A Chrome extension that lets users draw shapes with comments on Power BI Service reports, then export the result as PDF, PowerPoint, or Excel.

## Language

**Annotation**:
A single drawn shape with a comment, attached to one Page.
_Avoid_: marker, note, drawing, shape

**Tool**:
A way of drawing — rectangle, arrow, line, circle, or freehand. Each Tool knows how to render itself both live (while dragging) and from stored Annotation data.
_Avoid_: shape type, mode, brush

**Page**:
One Power BI report page, identified by a canonical key derived from report ID plus section hash. The unit Annotations attach to.
_Avoid_: tab, view, slide

**Report**:
A collection of Pages, identified by report ID. Has a defined Page order surfaced in Power BI's left navigation.
_Avoid_: dashboard, document

**PageStore**:
The module owning Page identity, Annotation persistence, and SPA navigation events. Hides Chrome storage, URL parsing, and Power BI sidebar DOM from the rest of the code.

**Capture**:
The flow that produces a screenshot of the current Power BI Page, gated by the user clicking the extension icon (uses `activeTab` permission).
_Avoid_: snapshot, screenshot grab

**Presentation**:
A logical export structure — one entry per Page, each with a title, captured screenshot, and numbered Comments list. Rendered to PPTX or PDF by a backend adapter.
_Avoid_: deck, slideshow, output

## Relationships

- A **Report** contains one or more **Pages**
- A **Page** contains zero or more **Annotations**
- An **Annotation** is drawn with exactly one **Tool**
- A **Capture** belongs to a **Page** and is consumed by **Presentation** rendering
- **PageStore** owns the set of all annotated **Pages** and notifies subscribers when the current **Page** changes

## Example dialogue

> **Dev:** "When the user navigates to a different **Page** in Power BI, what happens to the **Annotations** they already drew?"
> **Domain expert:** "Nothing — they stay attached to the original **Page**. **PageStore** notices the navigation, emits a page-change event, and the sidebar re-renders to show the new **Page**'s **Annotations** instead. The old ones are still in storage, keyed by their original **Page**."

## Flagged ambiguities

- "comment" was used to mean both the user-entered text on an **Annotation** and the list of those texts shown in an exported **Presentation** — resolved: "comment" refers to the text field; the rendered list is "the Presentation's Comments section."
- "page" was used to mean both a **Page** (Power BI report page) and a slide in the exported **Presentation** — resolved: exports are made of **Presentation** pages or slides; **Page** always refers to the Power BI source.
