(function () {
  const TAB_SELECTORS = [
    // Workspace report: bottom page tabs
    '.pagesNavigation button[role="tab"]',
    '.pagesNav button[role="tab"]',
    '[aria-label="Page navigation"] button[role="tab"]',
    'button[role="tab"]',
    '[class*="pageNavigator"] button',
    // App view: left-nav tree items / links (no bottom tabs)
    '[role="treeitem"]',
    '[role="tree"] a',
    'nav a',
    '[class*="navItem"]',
    'li[class*="page"] a',
  ];

  function cleanLabel(el) {
    const raw = el.getAttribute('aria-label') || el.textContent || '';
    return raw.replace(/[,\s]+(selected|active|current)$/i, '').trim();
  }

  function findNavElement(doc, { sectionHash, displayName }) {
    if (sectionHash) {
      const link = doc.querySelector(`a[href*="${sectionHash}"]`);
      if (link) return link;
    }
    if (displayName) {
      for (const selector of TAB_SELECTORS) {
        for (const el of doc.querySelectorAll(selector)) {
          if (cleanLabel(el) === displayName.trim()) return el;
        }
      }
    }
    return null;
  }

  const api = { findNavElement };
  if (typeof window !== 'undefined') window.PowerBIAnnotatorPageNavigator = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})();
