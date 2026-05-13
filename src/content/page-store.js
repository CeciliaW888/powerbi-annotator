function deriveKeyFromLocation(loc) {
  const reportIdFromPath = loc.pathname.match(/\/reports\/([^\/]+)/);
  let reportId = reportIdFromPath ? reportIdFromPath[1] : null;
  if (!reportId) {
    const params = new URLSearchParams(loc.search);
    reportId = params.get('reportId') || null;
  }

  const sectionFromPath = loc.pathname.match(/\/(ReportSection[a-zA-Z0-9]+)/);
  let sectionHash = sectionFromPath ? sectionFromPath[1] : null;
  if (!sectionHash) {
    const params = new URLSearchParams(loc.search);
    const pageName = params.get('pageName');
    if (pageName && pageName.startsWith('ReportSection')) {
      sectionHash = pageName;
    }
  }

  const key = reportId && sectionHash
    ? `${reportId}#${sectionHash}`
    : loc.pathname + loc.search;

  return { key, reportId, sectionHash };
}

function createPageStore({
  storage,
  locationProvider,
  displayNameResolver = () => 'Power BI Report',
  pageOrderResolver = () => [],
} = {}) {
  const cache = {};
  const pageChangeListeners = new Set();
  const dataChangeListeners = new Set();
  let lastKey = null;

  function deriveCurrent() {
    const loc = locationProvider();
    return deriveKeyFromLocation(loc);
  }

  function pageOf(key) {
    const { reportId, sectionHash } = (() => {
      const hashIdx = key.indexOf('#');
      if (hashIdx >= 0) {
        return { reportId: key.slice(0, hashIdx), sectionHash: key.slice(hashIdx + 1) };
      }
      return { reportId: null, sectionHash: null };
    })();
    return {
      key,
      reportId,
      sectionHash,
      displayName: displayNameResolver(sectionHash),
      annotations: cache[key] || [],
    };
  }

  function current() {
    const { key, reportId, sectionHash } = deriveCurrent();
    return {
      key,
      reportId,
      sectionHash,
      displayName: displayNameResolver(sectionHash),
      annotations: cache[key] || [],
    };
  }

  function list() {
    const order = pageOrderResolver() || [];
    const known = new Set(order);
    const ordered = order.filter(k => cache[k]);
    const extras = Object.keys(cache).filter(k => !known.has(k));
    return [...ordered, ...extras].map(pageOf);
  }

  function persist() {
    return new Promise((resolve) => {
      storage.set({ annotations: cache }, resolve);
    });
  }

  function notifyDataChange() {
    dataChangeListeners.forEach(cb => cb());
  }

  function saveAnnotations(annotations) {
    const { key } = deriveCurrent();
    cache[key] = annotations;
    persist();
    notifyDataChange();
  }

  function deleteAnnotations(key) {
    delete cache[key];
    persist();
    notifyDataChange();
  }

  function deleteAll() {
    Object.keys(cache).forEach(k => delete cache[k]);
    persist();
    notifyDataChange();
  }

  function onPageChange(cb) {
    pageChangeListeners.add(cb);
    return () => pageChangeListeners.delete(cb);
  }

  function onDataChange(cb) {
    dataChangeListeners.add(cb);
    return () => dataChangeListeners.delete(cb);
  }

  function checkPageChange() {
    const { key } = deriveCurrent();
    if (key !== lastKey) {
      lastKey = key;
      const page = current();
      pageChangeListeners.forEach(cb => cb(page));
    }
  }

  function init() {
    return new Promise((resolve) => {
      storage.get(['annotations'], (result) => {
        Object.keys(cache).forEach(k => delete cache[k]);
        Object.assign(cache, result.annotations || {});
        const { key } = deriveCurrent();
        lastKey = key;

        // Legacy migration: pathname-only keys → canonical key
        const loc = locationProvider();
        const legacyKey = loc.pathname;
        if (legacyKey !== key && cache[legacyKey] && !cache[key]) {
          cache[key] = cache[legacyKey];
          delete cache[legacyKey];
          persist().then(resolve);
          return;
        }
        resolve();
      });
    });
  }

  return {
    init,
    current,
    list,
    saveAnnotations,
    deleteAnnotations,
    deleteAll,
    onPageChange,
    onDataChange,
    checkPageChange,
    // Live reference to the internal cache. Provided to ease incremental
    // migration of legacy call sites that read annotations[pageKey] directly.
    // New code should use list() / current().
    _snapshot: () => cache,
  };
}

if (typeof window !== 'undefined') {
  window.PowerBIAnnotatorPageStore = { createPageStore };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { createPageStore };
}
