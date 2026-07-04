// Power BI Annotator - Content Script
// This script runs on Power BI pages and adds annotation functionality

let isAnnotationMode = false;
let annotations = [];
let currentAnnotation = null;
let startX, startY;
let sidebarOpen = false;
let currentDrawingTool = 'rectangle'; // rectangle, arrow, line, circle, freehand
let currentColor = '#0078d4';
let freehandPoints = [];
let annotationIdCounter = 0; // [Fix #11] Counter to avoid Date.now() collisions
let allAnnotationsCache = null; // Mirror of pageStore data for legacy call sites
let lastPageKey = null; // Tracked by the SPA navigation watcher (separate from pageStore's listener)
let lastReportId = null; // Track current report ID to detect report switches
let screenshotCache = {}; // { pageKey: dataUrl } - cached screenshots per page
let pageNameCache = {}; // { pageKey: displayName } - page names from Power BI embed API

// PageStore owns page identity, annotation persistence, and SPA navigation events.
// Constructed at init() time with concrete adapters; see CONTEXT.md for the seam.
let pageStore = null;

// --- Custom Modal Helpers (Fix #8: replace blocking prompt/alert/confirm) ---

/**
 * Show an informational modal (replaces alert).
 * Returns a promise that resolves when the user clicks OK.
 */
function showToast(message) {
  const existing = document.querySelector('.pbi-toast');
  if (existing) existing.remove();
  const toast = document.createElement('div');
  toast.className = 'pbi-toast';
  toast.textContent = message;
  document.body.appendChild(toast);
  requestAnimationFrame(() => toast.classList.add('visible'));
  setTimeout(() => {
    toast.classList.remove('visible');
    setTimeout(() => toast.remove(), 300);
  }, 2500);
}

function showModal(message) {
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.className = 'pbi-modal-overlay';
    overlay.innerHTML = `
      <div class="pbi-modal">
        <div class="pbi-modal-body"></div>
        <div class="pbi-modal-actions">
          <button class="pbi-modal-btn pbi-modal-btn-primary">OK</button>
        </div>
      </div>
    `;
    // Set text content separately to avoid XSS
    overlay.querySelector('.pbi-modal-body').textContent = message;
    document.body.appendChild(overlay);

    const okBtn = overlay.querySelector('.pbi-modal-btn-primary');
    okBtn.focus();
    okBtn.addEventListener('click', () => {
      overlay.remove();
      resolve();
    });
  });
}

/**
 * Show a confirmation modal (replaces confirm).
 * Returns a promise that resolves to true (OK) or false (Cancel).
 */
function showConfirm(message) {
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.className = 'pbi-modal-overlay';
    overlay.innerHTML = `
      <div class="pbi-modal">
        <div class="pbi-modal-body"></div>
        <div class="pbi-modal-actions">
          <button class="pbi-modal-btn pbi-modal-btn-cancel">Cancel</button>
          <button class="pbi-modal-btn pbi-modal-btn-primary">OK</button>
        </div>
      </div>
    `;
    overlay.querySelector('.pbi-modal-body').textContent = message;
    document.body.appendChild(overlay);

    const okBtn = overlay.querySelector('.pbi-modal-btn-primary');
    const cancelBtn = overlay.querySelector('.pbi-modal-btn-cancel');
    okBtn.focus();

    okBtn.addEventListener('click', () => {
      overlay.remove();
      resolve(true);
    });
    cancelBtn.addEventListener('click', () => {
      overlay.remove();
      resolve(false);
    });
  });
}

/**
 * Show a text input modal (replaces prompt).
 * Returns a promise that resolves to the entered text, or null if cancelled.
 */
function showPrompt(message) {
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.className = 'pbi-modal-overlay';
    overlay.innerHTML = `
      <div class="pbi-modal">
        <div class="pbi-modal-body"></div>
        <textarea class="pbi-modal-input" placeholder="Enter your comment..."></textarea>
        <div class="pbi-modal-actions">
          <button class="pbi-modal-btn pbi-modal-btn-cancel">Cancel</button>
          <button class="pbi-modal-btn pbi-modal-btn-primary">OK</button>
        </div>
      </div>
    `;
    overlay.querySelector('.pbi-modal-body').textContent = message;
    document.body.appendChild(overlay);

    const input = overlay.querySelector('.pbi-modal-input');
    const okBtn = overlay.querySelector('.pbi-modal-btn-primary');
    const cancelBtn = overlay.querySelector('.pbi-modal-btn-cancel');
    input.focus();

    okBtn.addEventListener('click', () => {
      const value = input.value;
      overlay.remove();
      resolve(value || null);
    });
    cancelBtn.addEventListener('click', () => {
      overlay.remove();
      resolve(null);
    });
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && e.ctrlKey) {
        const value = input.value;
        overlay.remove();
        resolve(value || null);
      }
    });
  });
}

// --- Utility Helpers ---

/**
 * Generate a unique annotation ID. [Fix #11]
 * Combines timestamp with a counter to avoid collisions.
 */
function generateAnnotationId() {
  return Date.now() * 100 + (annotationIdCounter++ % 100);
}

/**
 * Get the global starting number for annotations on the current page.
 * Counts all annotations on pages that come before this page in report order.
 */
function getGlobalStartNumber() {
  if (!allAnnotationsCache) return 0;
  
  const currentKey = getPageKey();
  const currentPageName = getPageName();
  const reportPageOrder = getReportPageOrder();
  const pages = getAnnotatedPages();
  
  let startNumber = 0;
  
  // Find current page in the ordered list and count all annotations before it
  for (const page of pages) {
    if (page.key === currentKey) {
      break;
    }
    startNumber += page.count;
  }
  
  return startNumber;
}

/**
 * Update the number badges on all annotation boxes to use global numbering. [Fix #3]
 * Called after deletion to keep page badges in sync with the sidebar.
 */
function renumberAnnotations() {
  const globalStart = getGlobalStartNumber();
  annotations.forEach((annotation, index) => {
    const box = document.querySelector(`.pbi-annotation-box[data-id="${annotation.id}"]`);
    if (box) {
      const badge = box.querySelector('.pbi-annotation-number');
      if (badge) {
        badge.textContent = globalStart + index + 1;
      }
    }
  });
}

/**
 * Check if any annotations are outside the visible viewport. [Fix #9]
 * Returns true if all annotations are visible on screen.
 */
function allAnnotationsInViewport() {
  const viewTop = window.scrollY;
  const viewBottom = viewTop + window.innerHeight;
  const viewLeft = window.scrollX;
  const viewRight = viewLeft + window.innerWidth;

  return annotations.every(a => {
    return a.x >= viewLeft && a.x + a.width <= viewRight &&
           a.y >= viewTop && a.y + a.height <= viewBottom;
  });
}

// --- Main Extension Logic ---

// Inject page-world script to access Power BI embed API (content scripts can't access page JS)
function injectPowerBIPageScript() {
  const script = document.createElement('script');
  script.src = chrome.runtime.getURL('src/content/powerbi-page-script.js');
  (document.head || document.documentElement).appendChild(script);
}

// Listen for messages from the injected page-world script
window.addEventListener('message', function(event) {
  if (event.source !== window) return;
  if (!event.data || !event.data.type) return;

  if (event.data.type === '__pbi_annotator_page_info__') {
    const { displayName, url } = event.data;
    if (displayName && url) {
      pageNameCache[url] = displayName;
      chrome.storage.local.set({ pageNameCache });
      renderPageList();
    }
  }

  // Handle navigation change from page-world pushState/replaceState override
  if (event.data.type === '__pbi_annotator_navigation__') {
    const currentKey = getPageKey();
    if (currentKey !== lastPageKey) {
      onPageChanged(lastPageKey, currentKey);
    }
  }
});

// Initialize the extension
function init() {
  pageStore = window.PowerBIAnnotatorPageStore.createPageStore({
    storage: chrome.storage.local,
    locationProvider: () => ({ pathname: window.location.pathname, search: window.location.search }),
    displayNameResolver: () => getPageName(),
    pageOrderResolver: () => getReportPageOrder(),
  });

  createSidebar();
  loadAnnotations();
  loadScreenshotCache();
  loadPageNameCache();
  setupEventListeners();
  lastPageKey = getPageKey();
  lastReportId = getReportId();
  startNavigationWatcher();
  watchCanvasLayout();
  injectPowerBIPageScript();
  console.log("Power BI Annotator initialized");
}

let canvasResizeObserver = null;
let repositionQueued = false;

function repositionAllAnnotations() {
  // rAF-coalesced: ResizeObserver can fire in bursts during PBI's own layout
  if (repositionQueued) return;
  repositionQueued = true;
  requestAnimationFrame(() => {
    repositionQueued = false;
    renderAnnotationsForCurrentPage();
  });
}

function watchCanvasLayout() {
  const canvas = getReportCanvas();
  if (!canvas) {
    // PBI renders the canvas late; retry until it exists
    setTimeout(watchCanvasLayout, 1000);
    return;
  }
  if (canvasResizeObserver) canvasResizeObserver.disconnect();
  canvasResizeObserver = new ResizeObserver(repositionAllAnnotations);
  canvasResizeObserver.observe(canvas);
  window.addEventListener('resize', repositionAllAnnotations);
}

// Poll for URL changes to detect SPA navigation (Power BI changes URL without page reload)
function startNavigationWatcher() {
  // Poll for URL changes (backup method)
  setInterval(() => {
    const currentKey = getPageKey();
    if (currentKey !== lastPageKey) {
      onPageChanged(lastPageKey, currentKey);
    }
  }, 300); // Reduced from 500ms to 300ms for faster detection

  // Also listen for browser navigation events
  window.addEventListener('popstate', () => {
    const currentKey = getPageKey();
    if (currentKey !== lastPageKey) {
      onPageChanged(lastPageKey, currentKey);
    }
  });

  window.addEventListener('hashchange', () => {
    const currentKey = getPageKey();
    if (currentKey !== lastPageKey) {
      onPageChanged(lastPageKey, currentKey);
    }
  });

  // Listen for pushState/replaceState (Power BI uses these for SPA navigation)
  const originalPushState = history.pushState;
  const originalReplaceState = history.replaceState;

  history.pushState = function(...args) {
    originalPushState.apply(this, args);
    setTimeout(() => {
      const currentKey = getPageKey();
      if (currentKey !== lastPageKey) {
        onPageChanged(lastPageKey, currentKey);
      }
    }, 100);
  };

  history.replaceState = function(...args) {
    originalReplaceState.apply(this, args);
    setTimeout(() => {
      const currentKey = getPageKey();
      if (currentKey !== lastPageKey) {
        onPageChanged(lastPageKey, currentKey);
      }
    }, 100);
  };
}

// Handle SPA page navigation: save current state, clear DOM, load new page's annotations
function onPageChanged(oldKey, newKey) {
  console.log('Page changed from', oldKey, 'to', newKey);
  
  // Check if we switched to a different report
  const currentReportId = getReportId();
  const reportChanged = lastReportId && currentReportId !== lastReportId;
  
  if (reportChanged) {
    console.log('Report changed from', lastReportId, 'to', currentReportId);
    // Turn off annotation mode when switching reports
    if (isAnnotationMode) {
      toggleAnnotationMode();
    }
  }
  
  lastReportId = currentReportId;
  
  // Immediately clear all annotation boxes from DOM (defensive - do this first)
  document.querySelectorAll('.pbi-annotation-box').forEach(box => box.remove());
  
  // Save current page's annotations under the old key (before URL changed)
  // Must use oldKey explicitly because getPageKey() now returns newKey
  if (allAnnotationsCache === null) {
    allAnnotationsCache = {};
  }
  allAnnotationsCache[oldKey] = annotations;
  chrome.storage.local.set({ annotations: allAnnotationsCache }, () => {
    if (chrome.runtime.lastError) {
      console.error("Failed to save annotations:", chrome.runtime.lastError);
    }
  });

  // Clear again (double-check - handle any boxes that might have been created during transition)
  document.querySelectorAll('.pbi-annotation-box').forEach(box => box.remove());

  // Load new page's annotations from in-memory cache
  if (allAnnotationsCache) {
    annotations = allAnnotationsCache[newKey] || [];
  } else {
    annotations = [];
  }

  // Power BI may still be re-laying-out the canvas right after nav; render
  // once now and again after the settle timeout below re-runs reposition.
  if (migrateLoadedAnnotations()) {
    pageStore.saveAnnotations(annotations);
  }
  renderAnnotationsForCurrentPage();

  lastPageKey = newKey;

  // Wait for Power BI to update the DOM before getting the page name
  // Power BI updates the page title/navigation after URL change, not instantly
  setTimeout(() => {
    repositionAllAnnotations();
    renderComments();
    renderPageList();
    // The canvas element can be replaced by PBI on page switch; re-observe the fresh node.
    watchCanvasLayout();
    // Cache screenshot after annotations are rendered on the new page
    if (annotations.length > 0) {
      setTimeout(() => cacheCurrentScreenshot(newKey), 600);
    }
  }, 400); // Give Power BI 400ms to update the DOM
}

// Request a silent screenshot from the background script (uses host_permissions, no user gesture needed)
function cacheCurrentScreenshot(pageKey) {
  const key = pageKey || getPageKey();
  // Don't cache if page has no annotations
  if (allAnnotationsCache && (!allAnnotationsCache[key] || allAnnotationsCache[key].length === 0)) {
    console.log('[Cache] Skipped - no annotations for:', key);
    return;
  }
  console.log('[Cache] Attempting screenshot for:', key);
  try {
    // Hide sidebar and toggle button so they don't appear in the screenshot
    const sidebar = document.getElementById('pbi-annotator-sidebar');
    const toggleBtn = document.getElementById('pbi-toggle-btn');
    if (sidebar) sidebar.style.display = 'none';
    if (toggleBtn) toggleBtn.style.display = 'none';

    // Wait for browser to repaint before capturing
    setTimeout(() => {
      chrome.runtime.sendMessage({ action: 'captureForCache' }, (response) => {
        // Restore sidebar and toggle button
        if (sidebar) sidebar.style.display = '';
        if (toggleBtn) toggleBtn.style.display = '';

        if (chrome.runtime.lastError) {
          console.log('[Cache] FAILED:', chrome.runtime.lastError.message);
          return;
        }
        if (response && response.screenshot) {
          screenshotCache[key] = response.screenshot;
          console.log('[Cache] SUCCESS for:', key, '| Total cached:', Object.keys(screenshotCache).length);
          // Limit to 20 most recent pages to manage storage size
          const keys = Object.keys(screenshotCache);
          if (keys.length > 20) {
            delete screenshotCache[keys[0]];
          }
          chrome.storage.local.set({ screenshotCache });
        } else {
          console.log('[Cache] FAILED for:', key, '| error:', response && response.error ? response.error : 'empty response');
        }
      });
    }, 200); // 200ms for browser to repaint
  } catch (e) {
    console.log('[Cache] Error:', e);
    // Restore UI elements on error
    const sidebar = document.getElementById('pbi-annotator-sidebar');
    const toggleBtn = document.getElementById('pbi-toggle-btn');
    if (sidebar) sidebar.style.display = '';
    if (toggleBtn) toggleBtn.style.display = '';
  }
}

// Load cached screenshots from storage on init
function loadScreenshotCache() {
  chrome.storage.local.get(['screenshotCache'], (result) => {
    if (!chrome.runtime.lastError && result.screenshotCache) {
      screenshotCache = result.screenshotCache;
    }
  });
}

function loadPageNameCache() {
  chrome.storage.local.get(['pageNameCache'], (result) => {
    if (!chrome.runtime.lastError && result.pageNameCache) {
      pageNameCache = result.pageNameCache;
    }
  });
}

// Resolver for pending screenshot capture (set during export, resolved by screenshotResult message)
let screenshotResolver = null;

// Listen for messages from background script
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  try {
    if (message.action === "toggleSidebar") {
      toggleSidebar();
      sendResponse({ success: true });
    } else if (message.action === "hideForCapture") {
      // Hide the instruction modal before the screenshot is taken
      const modal = document.querySelector('.pbi-modal-overlay');
      if (modal) modal.style.display = 'none';
      sendResponse({ success: true });
    } else if (message.action === "screenshotResult") {
      // Background captured the screenshot after the user clicked the extension icon
      if (screenshotResolver) {
        screenshotResolver(message);
        screenshotResolver = null;
      }
      sendResponse({ success: true });
    }
  } catch (error) {
    console.error('Error handling message:', error);
    sendResponse({ success: false, error: error.message });
  }
  return true; // Keep the message channel open for async response
});

// Create the sidebar UI
function createSidebar() {
  const sidebar = document.createElement("div");
  sidebar.id = "pbi-annotator-sidebar";
  sidebar.className = "pbi-sidebar";
  sidebar.innerHTML = `
    <div class="pbi-sidebar-header">
      <h3>Power BI Annotator<span id="pbi-total-count" class="pbi-count-badge"></span></h3>
      <button id="pbi-close-sidebar" class="pbi-btn-close">\u00d7</button>
    </div>
    <div class="pbi-sidebar-controls">
      <button id="pbi-toggle-annotate" class="pbi-btn pbi-btn-full">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4 12.5-12.5z"/></svg>
        <span class="pbi-btn-label">Start annotating</span>
      </button>
      <div id="pbi-drawing-toolbar" class="pbi-drawing-toolbar" style="display: none;">
        <div class="pbi-tool-label">Drawing Tool:</div>
        <button class="pbi-tool-btn active" data-tool="rectangle" title="Rectangle">
          <svg width="20" height="20" viewBox="0 0 20 20"><rect x="2" y="5" width="16" height="10" fill="none" stroke="currentColor" stroke-width="2"/></svg>
        </button>
        <button class="pbi-tool-btn" data-tool="arrow" title="Arrow">
          <svg width="20" height="20" viewBox="0 0 20 20"><path d="M2 10 L15 10 M15 10 L11 6 M15 10 L11 14" fill="none" stroke="currentColor" stroke-width="2"/></svg>
        </button>
        <button class="pbi-tool-btn" data-tool="circle" title="Circle">
          <svg width="20" height="20" viewBox="0 0 20 20"><circle cx="10" cy="10" r="7" fill="none" stroke="currentColor" stroke-width="2"/></svg>
        </button>
        <button class="pbi-tool-btn" data-tool="line" title="Line">
          <svg width="20" height="20" viewBox="0 0 20 20"><line x1="2" y1="15" x2="18" y2="5" stroke="currentColor" stroke-width="2"/></svg>
        </button>
        <button class="pbi-tool-btn" data-tool="freehand" title="Freehand">
          <svg width="20" height="20" viewBox="0 0 20 20"><path d="M2 15 Q 5 5, 10 10 T 18 8" fill="none" stroke="currentColor" stroke-width="2"/></svg>
        </button>
        <div class="pbi-swatch-row">
          <button class="pbi-swatch active" data-color="#0078d4" style="background:#0078d4" title="Blue"></button>
          <button class="pbi-swatch" data-color="#e81123" style="background:#e81123" title="Red"></button>
          <button class="pbi-swatch" data-color="#107c10" style="background:#107c10" title="Green"></button>
          <button class="pbi-swatch" data-color="#ffb900" style="background:#ffb900" title="Amber"></button>
          <button class="pbi-swatch" data-color="#5c2d91" style="background:#5c2d91" title="Purple"></button>
          <input type="color" id="pbi-color-picker" value="#0078d4" title="Custom color">
        </div>
      </div>
      <div class="pbi-button-row">
        <button id="pbi-export-pages" class="pbi-btn pbi-btn-primary">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
          <span>Export PDF / PPT</span>
        </button>
        <button id="pbi-export-annotations" class="pbi-btn pbi-btn-success">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/><line x1="15" y1="3" x2="15" y2="21"/></svg>
          <span>Export Excel</span>
        </button>
      </div>
      <div class="pbi-button-row">
        <button id="pbi-clear-page" class="pbi-btn pbi-btn-danger">
          Clear Page
        </button>
        <button id="pbi-clear-all" class="pbi-btn pbi-btn-danger">
          Clear All
        </button>
      </div>
    </div>
    <div class="pbi-page-section">
      <div class="pbi-page-header" id="pbi-page-header">
        <span class="pbi-page-indicator">
          <span class="pbi-page-icon">\ud83d\udcc4</span>
          <span id="pbi-current-page-name">Current Page</span>
        </span>
        <span id="pbi-page-count-badge" class="pbi-page-count-badge"></span>
      </div>
      <div class="pbi-page-list" id="pbi-page-list" style="display: none;"></div>
    </div>
    <div class="pbi-sidebar-content" id="pbi-comments-list">
      <div class="pbi-empty-state">
        <p><strong>No annotations yet</strong></p>
        <ol>
          <li>Click <em>Start annotating</em></li>
          <li>Drag on the report to draw a shape</li>
          <li>Type a comment — it appears here</li>
        </ol>
      </div>
    </div>
  `;

  document.body.appendChild(sidebar);

  // Create toggle button (slim edge tab, draggable)
  const toggleBtn = document.createElement("button");
  toggleBtn.id = "pbi-toggle-btn";
  toggleBtn.className = "pbi-toggle-btn";
  toggleBtn.innerHTML = "\ud83d\udcac";
  toggleBtn.title = "Toggle Comments Sidebar (drag to move)";
  document.body.appendChild(toggleBtn);

  // Restore saved position
  chrome.storage.local.get(['toggleBtnTop'], (result) => {
    if (!chrome.runtime.lastError && result.toggleBtnTop != null) {
      toggleBtn.style.top = result.toggleBtnTop + 'px';
      toggleBtn.style.transform = 'none';
    }
  });

  // Make toggle button draggable along the right edge
  let isDragging = false;
  let dragStartY = 0;
  let btnStartTop = 0;
  let hasDragged = false;

  toggleBtn.addEventListener('mousedown', (e) => {
    isDragging = true;
    hasDragged = false;
    dragStartY = e.clientY;
    btnStartTop = toggleBtn.getBoundingClientRect().top;
    toggleBtn.classList.add('dragging');
    e.preventDefault();
  });

  document.addEventListener('mousemove', (e) => {
    if (!isDragging) return;
    const dy = e.clientY - dragStartY;
    if (Math.abs(dy) > 3) hasDragged = true;
    let newTop = btnStartTop + dy;
    // Clamp to viewport
    newTop = Math.max(0, Math.min(window.innerHeight - toggleBtn.offsetHeight, newTop));
    toggleBtn.style.top = newTop + 'px';
    toggleBtn.style.transform = 'none';
  });

  document.addEventListener('mouseup', () => {
    if (!isDragging) return;
    isDragging = false;
    toggleBtn.classList.remove('dragging');
    // Persist position
    const top = parseInt(toggleBtn.style.top, 10);
    if (!isNaN(top)) {
      chrome.storage.local.set({ toggleBtnTop: top });
    }
  });

  // Store drag state on the element so setupEventListeners can check it
  toggleBtn._getDragState = () => hasDragged;
}

// Setup event listeners
function setupEventListeners() {
  // Toggle sidebar (skip if user just finished dragging the button)
  const toggleBtnEl = document.getElementById("pbi-toggle-btn");
  toggleBtnEl.addEventListener("click", (e) => {
    if (toggleBtnEl._getDragState && toggleBtnEl._getDragState()) return;
    toggleSidebar();
  });
  document
    .getElementById("pbi-close-sidebar")
    .addEventListener("click", toggleSidebar);

  // Toggle annotation mode
  document
    .getElementById("pbi-toggle-annotate")
    .addEventListener("click", toggleAnnotationMode);

  // Clear page annotations
  document
    .getElementById("pbi-clear-page")
    .addEventListener("click", clearCurrentPageAnnotations);

  // Clear all annotations
  document
    .getElementById("pbi-clear-all")
    .addEventListener("click", clearAllAnnotations);

  // Export Pages (PDF/PPT with screenshots)
  document
    .getElementById("pbi-export-pages")
    .addEventListener("click", exportPages);

  // Export Annotations (Excel/CSV)
  document
    .getElementById("pbi-export-annotations")
    .addEventListener("click", exportAnnotations);

  // Drawing tool buttons
  document.querySelectorAll(".pbi-tool-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      const tool = e.currentTarget.dataset.tool;
      selectDrawingTool(tool);
    });
  });

  // Color swatches
  document.querySelectorAll('.pbi-swatch').forEach((swatch) => {
    swatch.addEventListener('click', () => {
      currentColor = swatch.dataset.color;
      document.getElementById('pbi-color-picker').value = currentColor;
      document.querySelectorAll('.pbi-swatch').forEach((s) => s.classList.toggle('active', s === swatch));
    });
  });

  // Color picker
  document.getElementById("pbi-color-picker").addEventListener("change", (e) => {
    currentColor = e.target.value;
    document.querySelectorAll('.pbi-swatch').forEach((s) => s.classList.remove('active'));
  });

  // Page list toggle
  document.getElementById("pbi-page-header").addEventListener("click", () => {
    const pageList = document.getElementById("pbi-page-list");
    const isVisible = pageList.style.display !== "none";
    pageList.style.display = isVisible ? "none" : "block";
    renderPageList();
  });

  // Mouse events for drawing annotations
  document.addEventListener("mousedown", handleMouseDown);
  document.addEventListener("mousemove", handleMouseMove);
  document.addEventListener("mouseup", handleMouseUp);

  // Prevent sidebar clicks from triggering annotations
  document
    .getElementById("pbi-annotator-sidebar")
    .addEventListener("mousedown", (e) => {
      e.stopPropagation();
    });
}

// Toggle sidebar visibility
function toggleSidebar() {
  const sidebar = document.getElementById("pbi-annotator-sidebar");
  sidebarOpen = !sidebarOpen;

  if (sidebarOpen) {
    sidebar.classList.add("open");
  } else {
    sidebar.classList.remove("open");
  }
}

// Toggle annotation mode
function toggleAnnotationMode() {
  isAnnotationMode = !isAnnotationMode;
  // While drawing, existing boxes must not swallow mouse events — otherwise
  // drawing over a previous annotation fires its view-comment click handler
  // and collides with the new-comment prompt (pointer-events rule in CSS).
  document.body.classList.toggle("pbi-annotating", isAnnotationMode);
  const btn = document.getElementById("pbi-toggle-annotate");
  const toolbar = document.getElementById("pbi-drawing-toolbar");

  const label = btn.querySelector(".pbi-btn-label");
  if (isAnnotationMode) {
    if (label) label.textContent = "Stop annotating";
    btn.classList.add("active");
    toolbar.style.display = "flex";
    document.body.style.cursor = "crosshair";
    // Hide sidebar so it doesn't obstruct the annotation area
    const sidebar = document.getElementById("pbi-annotator-sidebar");
    if (sidebar && sidebarOpen) {
      sidebarOpen = false;
      sidebar.classList.remove("open");
    }
  } else {
    if (label) label.textContent = "Start annotating";
    btn.classList.remove("active");
    toolbar.style.display = "none";
    document.body.style.cursor = "default";
  }
}

// Select drawing tool
function selectDrawingTool(tool) {
  currentDrawingTool = tool;

  // Update active state
  document.querySelectorAll(".pbi-tool-btn").forEach((btn) => {
    btn.classList.remove("active");
  });
  document.querySelector(`[data-tool="${tool}"]`).classList.add("active");
}

// Handle mouse down - start drawing annotation
function handleMouseDown(e) {
  if (!isAnnotationMode) return;
  if (e.target.closest("#pbi-annotator-sidebar")) return;
  if (e.target.closest("#pbi-drawing-toolbar")) return;
  if (e.target.closest(".pbi-modal-overlay")) return;
  
  // Allow clicking on Power BI navigation and UI controls without creating annotations
  if (e.target.closest('button[role="tab"]')) return; // Page navigation tabs
  if (e.target.closest('.navigationPane')) return; // Navigation panel
  if (e.target.closest('.pagesNav')) return; // Pages navigation
  if (e.target.closest('[class*="pageNavigator"]')) return; // Page navigator
  if (e.target.closest('button')) return; // Any button (filters, slicers, etc.)
  if (e.target.closest('a')) return; // Any link
  if (e.target.closest('[role="button"]')) return; // Elements acting as buttons
  if (e.target.closest('input, select, textarea')) return; // Form controls
  if (e.target.closest('[class*="slicer"]')) return; // Slicers
  if (e.target.closest('[class*="filter"]')) return; // Filters

  startX = e.pageX;
  startY = e.pageY;

  // Create temporary annotation element
  currentAnnotation = document.createElement("div");
  currentAnnotation.className = "pbi-annotation-box pbi-drawing";
  currentAnnotation.style.left = startX + "px";
  currentAnnotation.style.top = startY + "px";
  currentAnnotation.style.borderColor = currentColor;

  if (currentDrawingTool === 'freehand') {
    freehandPoints = [{ x: startX, y: startY }];
  } else {
    currentAnnotation.style.width = "0px";
    currentAnnotation.style.height = "0px";
  }

  document.body.appendChild(currentAnnotation);
}

// Handle mouse move - resize annotation
function handleMouseMove(e) {
  if (!currentAnnotation) return;
  const Tools = window.PowerBIAnnotatorTools;

  if (currentDrawingTool === 'freehand') {
    freehandPoints.push({ x: e.pageX, y: e.pageY });
  }

  const geometry = Tools.computeGeometry(
    currentDrawingTool,
    { x: startX, y: startY },
    { x: e.pageX, y: e.pageY },
    currentDrawingTool === 'freehand' ? freehandPoints : null
  );

  currentAnnotation.style.left = geometry.x + 'px';
  currentAnnotation.style.top = geometry.y + 'px';
  currentAnnotation.style.width = geometry.width + 'px';
  currentAnnotation.style.height = geometry.height + 'px';

  updateAnnotationVisual(currentAnnotation, geometry);
}

function updateAnnotationVisual(element, geometry) {
  const existingSvg = element.querySelector('svg');
  if (existingSvg) existingSvg.remove();

  const tool = window.PowerBIAnnotatorTools[currentDrawingTool];
  if (!tool) return;

  const svg = tool.render(geometry, currentColor);
  if (svg) element.appendChild(svg);
}

// Handle mouse up - finish annotation and prompt for comment
// [Fix #1, #2, #8, #11] - async for custom prompt, stores direction, uses appendChild, unique IDs
async function handleMouseUp(e) {
  if (!currentAnnotation) return;

  // Capture references before clearing (prevents interference during async prompt)
  const finishedAnnotation = currentAnnotation;
  const drawStartX = startX;
  const drawStartY = startY;
  const drawEndX = e.pageX;
  const drawEndY = e.pageY;
  const toolUsed = currentDrawingTool;
  const colorUsed = currentColor;
  const capturedFreehandPoints = currentDrawingTool === 'freehand' ? [...freehandPoints] : null;
  currentAnnotation = null;

  const rect = finishedAnnotation.getBoundingClientRect();

  // Ignore very small annotations (accidental clicks)
  if (rect.width < 10 || rect.height < 10) {
    finishedAnnotation.remove();
    return;
  }

  // Prompt for comment using custom modal
  const comment = await showPrompt("Enter your comment for this annotation:");

  if (comment && comment.trim()) {
    let annotation = {
      id: generateAnnotationId(),
      x: parseInt(finishedAnnotation.style.left),
      y: parseInt(finishedAnnotation.style.top),
      width: rect.width,
      height: rect.height,
      comment: comment.trim(),
      timestamp: new Date().toISOString(),
      url: window.location.href,
      pageName: getPageName(),
      tool: toolUsed,
      color: colorUsed,
      freehandPath: capturedFreehandPoints,
      startPoint: { x: drawStartX, y: drawStartY },
      endPoint: { x: drawEndX, y: drawEndY },
    };

    // Anchor to the report canvas so the shape survives layout changes
    // (App view, window resize, sidebar toggles). Falls back to legacy
    // absolute coords when no canvas is found (e.g. test-page.html edge cases).
    const reportCanvas = getReportCanvas();
    if (reportCanvas) {
      const canvasRect = window.PowerBIAnnotatorCoords.getCanvasPageRect(reportCanvas, window);
      annotation = window.PowerBIAnnotatorCoords.annotationToRelative(annotation, canvasRect);
    }

    annotations.push(annotation);
    saveAnnotations();

    // Update the annotation box with ID
    finishedAnnotation.classList.remove("pbi-drawing");
    finishedAnnotation.dataset.id = annotation.id;

    // [Fix #1] Use appendChild instead of innerHTML += to preserve SVG elements
    if (toolUsed === 'rectangle') {
      finishedAnnotation.innerHTML = '';
    }
    const badge = document.createElement('div');
    badge.className = 'pbi-annotation-number';
    const globalStart = getGlobalStartNumber();
    badge.textContent = globalStart + annotations.length;
    finishedAnnotation.appendChild(badge);

    // Add click handler to show comment
    finishedAnnotation.addEventListener("click", (e) => {
      e.stopPropagation();
      showAnnotationComment(annotation.id);
    });

    renderComments();
    renderPageList();

    // Cache screenshot after creating an annotation (page is visible and has content)
    cacheCurrentScreenshot();
  } else {
    finishedAnnotation.remove();
  }
}

// Show annotation comment in custom modal [Fix #8]
function showAnnotationComment(id) {
  const annotation = annotations.find((a) => a.id === id);
  if (annotation) {
    const globalStart = getGlobalStartNumber();
    const localIndex = annotations.indexOf(annotation);
    showModal(
      `Comment #${globalStart + localIndex + 1}:\n\n${annotation.comment}`,
    );
  }
}

// Render comments in sidebar
function renderComments() {
  const commentsList = document.getElementById("pbi-comments-list");

  const countBadge = document.getElementById('pbi-total-count');
  if (countBadge) countBadge.textContent = annotations.length || '';

  if (annotations.length === 0) {
    commentsList.innerHTML =
      '<div class="pbi-empty-state"><p><strong>No annotations yet</strong></p><ol><li>Click <em>Start annotating</em></li><li>Drag on the report to draw a shape</li><li>Type a comment — it appears here</li></ol></div>';
    return;
  }

  const globalStart = getGlobalStartNumber();
  commentsList.innerHTML = annotations
    .map(
      (annotation, index) => `
    <div class="pbi-comment-item" data-id="${annotation.id}">
      <div class="pbi-comment-header">
        <span class="pbi-comment-number" style="background:${annotation.color || '#0078d4'}">#${globalStart + index + 1}</span>
        <span class="pbi-comment-time">${formatTime(annotation.timestamp)}</span>
      </div>
      <div class="pbi-comment-text">${escapeHtml(annotation.comment)}</div>
      <div class="pbi-comment-actions">
        <button class="pbi-btn-small pbi-btn-highlight" data-id="${annotation.id}">
          Highlight
        </button>
        <button class="pbi-btn-small pbi-btn-delete" data-id="${annotation.id}">
          Delete
        </button>
      </div>
    </div>
  `,
    )
    .join("");

  // Add event listeners for highlight and delete buttons
  commentsList.querySelectorAll(".pbi-btn-highlight").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      highlightAnnotation(parseInt(e.target.dataset.id));
    });
  });

  commentsList.querySelectorAll(".pbi-btn-delete").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      deleteAnnotation(parseInt(e.target.dataset.id));
    });
  });
}

// Get the report page order from Power BI navigation
function getReportPageOrder() {
  const pageOrder = [];
  
  // Try to find all page navigation buttons in order
  const navigationSelectors = [
    'button[role="tab"][aria-label]',
    '.navigationPane button[aria-label]',
    'button[aria-label*="Page"]',
    '.pagesNav button',
    'div[role="tablist"] button'
  ];
  
  for (const selector of navigationSelectors) {
    const buttons = document.querySelectorAll(selector);
    if (buttons.length > 0) {
      buttons.forEach(btn => {
        const pageName = btn.getAttribute('aria-label') || btn.getAttribute('title') || btn.textContent?.trim();
        if (pageName && pageName !== 'Page navigation' && !pageName.includes('navigation')) {
          // Clean up the name
          const cleanName = pageName
            .replace(/[,\s]+selected$/i, '')
            .replace(/[,\s]+active$/i, '')
            .replace(/^Page\s+/i, '')
            .trim();
          if (cleanName && !pageOrder.includes(cleanName)) {
            pageOrder.push(cleanName);
          }
        }
      });
      if (pageOrder.length > 0) break;
    }
  }
  
  return pageOrder;
}

// Get list of all pages that have annotations in the current report
function getAnnotatedPages() {
  if (!allAnnotationsCache) return [];
  const currentKey = getPageKey();
  const currentReportId = getReportId();
  const reportPageOrder = getReportPageOrder();
  
  const pages = Object.keys(allAnnotationsCache)
    .filter(key => {
      if (allAnnotationsCache[key].length === 0) return false;
      // Filter to only show pages from current report
      // Extract report ID from the stored page URL
      const firstAnnotation = allAnnotationsCache[key][0];
      if (firstAnnotation && firstAnnotation.url) {
        try {
          const url = new URL(firstAnnotation.url);
          const storedPath = url.pathname;
          const storedReportMatch = storedPath.match(/\/reports\/([^\/]+)/);
          const storedReportId = storedReportMatch
            ? storedReportMatch[1]
            : storedPath.split('/').filter(p => p).join('_');
          return storedReportId === currentReportId;
        } catch (e) {
          // If URL parsing fails, include the page
          return true;
        }
      }
      return true;
    })
    .map(key => {
      const pageAnnotations = allAnnotationsCache[key];
      // Get page name from stored annotation data (more reliable than URL parsing)
      let name = key;
      if (pageAnnotations.length > 0) {
        // Use stored pageName from first annotation if available
        if (pageAnnotations[0].pageName) {
          name = pageAnnotations[0].pageName;
        } else if (key === currentKey) {
          // Fallback: use getPageName() only for current page
          name = getPageName();
        } else {
          // Fallback: try to extract from URL
          try {
            const url = new URL(pageAnnotations[0].url || '');
            const parts = url.pathname.split('/').filter(p => p);
            name = parts.length > 0 ? decodeURIComponent(parts[parts.length - 1]) : key;
          } catch (e) {
            name = key;
          }
        }
      }
      return {
        key,
        name,
        count: pageAnnotations.length,
        isCurrent: key === currentKey,
        hasScreenshot: !!screenshotCache[key]
      };
    });
  
  // Sort pages to match report order
  if (reportPageOrder.length > 0) {
    pages.sort((a, b) => {
      const indexA = reportPageOrder.indexOf(a.name);
      const indexB = reportPageOrder.indexOf(b.name);
      
      // If both pages are in the report order, sort by that order
      if (indexA !== -1 && indexB !== -1) {
        return indexA - indexB;
      }
      // If only one is in the report order, it comes first
      if (indexA !== -1) return -1;
      if (indexB !== -1) return 1;
      // If neither is in the report order, sort alphabetically
      return a.name.localeCompare(b.name);
    });
  } else {
    // Fallback: sort alphabetically if we can't detect report order
    pages.sort((a, b) => a.name.localeCompare(b.name));
  }
  
  return pages;
}

// Render the page indicator and collapsible page list in sidebar
function renderPageList() {
  const pageNameEl = document.getElementById('pbi-current-page-name');
  const countBadge = document.getElementById('pbi-page-count-badge');
  const pageList = document.getElementById('pbi-page-list');

  if (!pageNameEl || !countBadge || !pageList) return;

  const pages = getAnnotatedPages();
  const currentPageName = getPageName();

  // Update current page indicator
  pageNameEl.textContent = currentPageName;

  // Update page count badge
  if (pages.length > 1) {
    countBadge.textContent = pages.length + ' pages';
    countBadge.style.display = '';
  } else {
    countBadge.textContent = '';
    countBadge.style.display = 'none';
  }

  // Update page list (only if visible)
  if (pageList.style.display === 'none') return;

  if (pages.length === 0) {
    pageList.innerHTML = '<p class="pbi-page-empty">No pages annotated yet.</p>';
    return;
  }

  pageList.innerHTML = pages.map(page => `
    <div class="pbi-page-item ${page.isCurrent ? 'pbi-page-current' : ''}">
      <div class="pbi-page-info">
        <span class="pbi-page-name">${escapeHtml(page.name)}</span>
        <span class="pbi-page-meta">${page.count} comment${page.count !== 1 ? 's' : ''}${page.hasScreenshot ? ' \u2022 \ud83d\udcf8' : ''}</span>
      </div>
      ${page.isCurrent ? '<span class="pbi-page-badge-current">Current</span>' : ''}
    </div>
  `).join('');
}

// Highlight annotation on page
function highlightAnnotation(id) {
  const annotationBox = document.querySelector(
    `.pbi-annotation-box[data-id="${id}"]`,
  );
  if (annotationBox) {
    // Scroll to annotation
    annotationBox.scrollIntoView({ behavior: "smooth", block: "center" });

    // Flash highlight effect
    annotationBox.classList.add("pbi-highlight");
    setTimeout(() => {
      annotationBox.classList.remove("pbi-highlight");
    }, 2000);
  }
}

// Delete annotation [Fix #3, #8] - async for custom confirm, renumbers badges
async function deleteAnnotation(id) {
  const confirmed = await showConfirm("Delete this comment?");
  if (!confirmed) return;

  // Remove from array
  annotations = annotations.filter((a) => a.id !== id);

  // Remove from DOM
  const annotationBox = document.querySelector(
    `.pbi-annotation-box[data-id="${id}"]`,
  );
  if (annotationBox) {
    annotationBox.remove();
  }

  saveAnnotations();
  renumberAnnotations();
  renderComments();
  renderPageList();
}

// Clear current page annotations
async function clearCurrentPageAnnotations() {
  const pageName = getPageName();
  const count = annotations.length;
  
  if (count === 0) {
    await showModal("No comments on current page.");
    return;
  }

  const confirmed = await showConfirm(
    `Delete all ${count} comment${count > 1 ? 's' : ''} from "${pageName}"?\n\nThis cannot be undone.`
  );
  if (!confirmed) return;

  annotations = [];
  document
    .querySelectorAll(".pbi-annotation-box")
    .forEach((box) => box.remove());
  saveAnnotations();
  renderComments();
  renderPageList();
  showToast('Page annotations cleared');
}

// Clear all annotations across all pages [Fix #8] - async for custom confirm
async function clearAllAnnotations() {
  // Count total annotations across all pages
  const pages = getAnnotatedPages();
  const totalCount = pages.reduce((sum, p) => sum + p.count, 0);
  
  if (totalCount === 0) {
    await showModal("No comments to delete.");
    return;
  }

  const confirmed = await showConfirm(
    `Delete ALL comments from ALL pages (${totalCount} total across ${pages.length} page${pages.length > 1 ? 's' : ''})?\n\nThis cannot be undone.`
  );
  if (!confirmed) return;

  // Clear current page UI
  annotations = [];
  document
    .querySelectorAll(".pbi-annotation-box")
    .forEach((box) => box.remove());
  
  // Clear all pages from storage via PageStore (mirror stays in sync because
  // _snapshot() returns the same object reference; deleteAll mutates in place).
  if (pageStore) pageStore.deleteAll();

  renderComments();
  renderPageList();
  showToast('All annotations cleared');
}

// Show scope selection dialog when multiple pages have annotations
function showScopeDialog() {
  const pages = getAnnotatedPages();
  if (pages.length <= 1) return Promise.resolve('current'); // Only one page, no need to ask

  const totalCount = pages.reduce((sum, p) => sum + p.count, 0);
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.className = 'pbi-modal-overlay';
    overlay.innerHTML = `
      <div class="pbi-modal">
        <div class="pbi-modal-body">Export scope</div>
        <p style="color:#666;font-size:13px;margin:0 0 20px 0;">
          You have annotations on ${pages.length} pages (${totalCount} total comments).
        </p>
        <div class="pbi-modal-actions" style="flex-direction:column;gap:8px;">
          <button class="pbi-modal-btn pbi-modal-btn-primary" style="width:100%;text-align:center;" data-scope="all">
            Export All Pages (${totalCount} comments)
          </button>
          <button class="pbi-modal-btn" style="width:100%;text-align:center;background:#f0f7ff;color:#0078d4;" data-scope="current">
            Current Page Only (${annotations.length} comments)
          </button>
          <button class="pbi-modal-btn pbi-modal-btn-cancel" style="width:100%;text-align:center;" data-scope="cancel">
            Cancel
          </button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    overlay.addEventListener('click', (e) => {
      const scope = e.target.dataset.scope;
      if (scope) {
        overlay.remove();
        resolve(scope === 'cancel' ? null : scope);
      }
    });
  });
}

// Build Excel data for a set of pages (used by both single and multi-page export)
function buildExcelData(pages) {
  const headers = ["No", "Page Name", "URL", "Date", "Comment"];
  const data = [headers];
  let globalNumber = 1;

  for (const page of pages) {
    // Use URL from first annotation (most accurate for that page's state)
    let pageUrl = page.url || page.key || '';
    if (page.annotations && page.annotations.length > 0 && page.annotations[0].url) {
      pageUrl = page.annotations[0].url;
    }

    for (const annotation of page.annotations) {
      const date = new Date(annotation.timestamp);
      data.push([
        globalNumber++,
        page.name,
        pageUrl,
        date.toLocaleDateString(),
        annotation.comment,
      ]);
    }
  }

  return data;
}

// Export annotations to Excel (.xlsx format) [Fix #8, #10]
async function exportAnnotations() {
  if (annotations.length === 0) {
    await showModal("No comments to export. Create some annotations first!");
    return;
  }

  // Check if multiple pages have annotations
  const scope = await showScopeDialog();
  if (!scope) return; // User cancelled

  let pages;
  if (scope === 'all') {
    pages = getAnnotatedPages().map(p => ({
      name: p.name,
      key: p.key,
      annotations: allAnnotationsCache[p.key] || []
    }));
  } else {
    pages = [{ name: getPageName(), key: getPageKey(), annotations }];
  }

  const totalCount = pages.reduce((sum, p) => sum + p.annotations.length, 0);
  const excelData = buildExcelData(pages);

  // Create workbook and worksheet using SheetJS
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(excelData);

  ws['!cols'] = [
    { wch: 6 },  // No
    { wch: 30 }, // Page Name
    { wch: 80 }, // URL
    { wch: 12 }, // Date
    { wch: 60 }  // Comment
  ];

  XLSX.utils.book_append_sheet(wb, ws, "Annotations");

  const now = new Date();
  const suffix = scope === 'all' ? 'AllPages' : 'Comments';
  const filename = `PowerBI_${suffix}_${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}-${String(now.getMinutes()).padStart(2, "0")}.xlsx`;

  XLSX.writeFile(wb, filename);

  const pageLabel = scope === 'all' ? ` across ${pages.length} page(s)` : '';
  await showModal(`Exported ${totalCount} comment(s)${pageLabel} to ${filename}`);
}

// Export pages with screenshots (PDF or PPT) [Fix #8, #9]
async function exportPages() {
  if (annotations.length === 0) {
    await showModal('No comments to export. Create some annotations first!');
    return;
  }

  // [Fix #9] Warn if annotations are outside the visible viewport
  if (!allAnnotationsInViewport()) {
    const proceed = await showConfirm(
      'Some annotations are outside the visible area and may not appear in the screenshot.\n\nScroll to make all annotations visible before exporting, or click OK to continue anyway.'
    );
    if (!proceed) return;
  }

  // Check if multiple pages have annotations
  const scope = await showScopeDialog();
  if (!scope) return; // User cancelled

  // Ask user for format
  const format = await showFormatDialog();
  if (!format) return; // User cancelled

  if (scope === 'all') {
    await generateMultiPagePresentation(format);
  } else {
    await generatePresentation(format);
  }
}

// Show format selection dialog
function showFormatDialog() {
  return new Promise((resolve) => {
    const dialog = document.createElement('div');
    dialog.className = 'pbi-format-dialog';
    dialog.innerHTML = `
      <div class="pbi-format-dialog-overlay"></div>
      <div class="pbi-format-dialog-content">
        <h3>Export Format</h3>
        <p>Choose the format for exporting pages with screenshots:</p>
        <div class="pbi-format-buttons">
          <button class="pbi-format-btn" data-format="pdf">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
              <polyline points="14 2 14 8 20 8"></polyline>
              <text x="7" y="17" font-size="6" fill="currentColor">PDF</text>
            </svg>
            <span>PDF</span>
          </button>
          <button class="pbi-format-btn" data-format="ppt">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
              <polyline points="14 2 14 8 20 8"></polyline>
              <text x="7" y="17" font-size="6" fill="currentColor">PPT</text>
            </svg>
            <span>PowerPoint</span>
          </button>
        </div>
        <button class="pbi-format-cancel">Cancel</button>
      </div>
    `;

    document.body.appendChild(dialog);

    // Handle button clicks
    dialog.querySelectorAll('.pbi-format-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const format = btn.dataset.format;
        dialog.remove();
        resolve(format);
      });
    });

    dialog.querySelector('.pbi-format-cancel').addEventListener('click', () => {
      dialog.remove();
      resolve(null);
    });

    // Close on overlay click
    dialog.querySelector('.pbi-format-dialog-overlay').addEventListener('click', () => {
      dialog.remove();
      resolve(null);
    });
  });
}

/**
 * Show a modal asking the user to click the extension icon, and wait for
 * the background script to send back the screenshot result.
 * Returns the screenshot result object, or null if the user cancels.
 */
function waitForScreenshotOrCancel() {
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.className = 'pbi-modal-overlay';
    overlay.innerHTML = `
      <div class="pbi-modal">
        <div class="pbi-modal-body" style="text-align: center;">
          <div style="font-size: 48px; margin-bottom: 16px;">📸</div>
          <p style="font-size: 16px; margin-bottom: 8px;"><strong>Click the Power BI Annotator icon in your browser toolbar to capture the screenshot.</strong></p>
          <p style="color: #666; font-size: 13px;">The extension icon should show a 📸 badge. It's in the top-right area of your browser, near the address bar.</p>
        </div>
        <div class="pbi-modal-actions">
          <button class="pbi-modal-btn pbi-modal-btn-cancel">Cancel</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    // When the screenshot comes back from background, remove modal and resolve
    screenshotResolver = (result) => {
      if (overlay.parentNode) overlay.remove();
      resolve(result);
    };

    // If user cancels, tell background to clear pending state
    overlay.querySelector('.pbi-modal-btn-cancel').addEventListener('click', () => {
      screenshotResolver = null;
      overlay.remove();
      try {
        chrome.runtime.sendMessage({ action: 'cancelCapture' }, (response) => {
          if (chrome.runtime.lastError) {
            console.log('Could not cancel capture:', chrome.runtime.lastError.message);
          }
        });
      } catch (error) {
        console.log('Extension context error:', error);
      }
      resolve(null);
    });
  });
}

// Generate presentation with screenshots [Fix #6, #8, #10]
async function generatePresentation(format) {
  // Hide sidebar during screenshot capture
  const sidebar = document.getElementById('pbi-annotator-sidebar');
  const toggleBtn = document.getElementById('pbi-toggle-btn');
  const sidebarWasOpen = sidebar.classList.contains('open');

  sidebar.style.display = 'none';
  toggleBtn.style.display = 'none';

  // Make sure all annotations are visible
  document.querySelectorAll('.pbi-annotation-box').forEach(box => {
    box.style.opacity = '1';
    box.style.zIndex = '';
  });

  // Get the Power BI report canvas area
  const reportCanvas = getReportCanvas();
  if (!reportCanvas) {
    sidebar.style.display = '';
    toggleBtn.style.display = '';
    if (sidebarWasOpen) sidebar.classList.add('open');
    await showModal('Could not find Power BI report canvas. Make sure you are on a report page.');
    return;
  }

  // Wait for rendering
  await new Promise(resolve => setTimeout(resolve, 300));

  // Tell background to wait for an icon click to capture the screenshot
  try {
    chrome.runtime.sendMessage({ action: 'prepareCapture' }, (response) => {
      if (chrome.runtime.lastError) {
        console.error('Could not prepare capture:', chrome.runtime.lastError.message);
      }
    });
  } catch (error) {
    console.error('Extension communication error:', error);
    sidebar.style.display = '';
    toggleBtn.style.display = '';
    if (sidebarWasOpen) sidebar.classList.add('open');
    await showModal('Unable to communicate with extension background. Try refreshing the page.');
    return;
  }

  // Show modal asking user to click the extension icon, wait for result or cancel
  let screenshot = null;
  try {
    const result = await waitForScreenshotOrCancel();
    if (result === null) {
      // User cancelled - restore sidebar and abort
      sidebar.style.display = '';
      toggleBtn.style.display = '';
      if (sidebarWasOpen) sidebar.classList.add('open');
      return;
    }
    if (result.screenshot) {
      // Crop screenshot to only the report canvas area
      screenshot = await cropScreenshotToCanvas(result.screenshot, reportCanvas);
    } else if (result.error) {
      console.error('Screenshot capture error:', result.error);
    }
  } catch (error) {
    console.error('Failed to capture screenshot:', error);
  }

  // Restore sidebar
  sidebar.style.display = '';
  toggleBtn.style.display = '';
  if (sidebarWasOpen) {
    sidebar.classList.add('open');
  }

  // Prepare all comments data
  const comments = annotations.map((annotation, index) => ({
    number: index + 1,
    comment: annotation.comment,
    date: new Date(annotation.timestamp).toLocaleDateString(),
    tool: annotation.tool || 'rectangle',
    color: annotation.color || '#0078d4'
  }));

  const pageName = getPageName();

  // Fork: PPT format generates a real .pptx file
  if (format === 'ppt') {
    await generatePptx(screenshot, comments, pageName);
    showToast('Export ready — check your downloads');
    return;
  }

  // PDF format: generate real .pdf file using jsPDF
  await generatePdf(screenshot, comments, pageName);
  showToast('Export ready — check your downloads');
}

/**
 * Generate a multi-page presentation/PDF that includes all annotated pages.
 * Uses cached screenshots for other pages and captures a fresh screenshot for the current page.
 */
function captureSilently() {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({ action: 'captureForCache' }, (response) => {
      if (chrome.runtime.lastError || !response || !response.screenshot) {
        resolve(null);
        return;
      }
      resolve(response.screenshot);
    });
  });
}

// Silent first; icon-click modal once as fallback. After the first grant,
// silent capture works for the rest of the wizard loop.
async function captureVisiblePage() {
  const silent = await captureSilently();
  if (silent) return silent;
  chrome.runtime.sendMessage({ action: 'prepareCapture' }, () => { void chrome.runtime.lastError; });
  const result = await waitForScreenshotOrCancel();
  return result && result.screenshot ? result.screenshot : null;
}

function showExportProgress(pageNames) {
  const overlay = document.createElement('div');
  overlay.className = 'pbi-modal-overlay';
  overlay.id = 'pbi-export-progress';
  overlay.innerHTML = `
    <div class="pbi-modal">
      <div class="pbi-modal-header"><h3>Exporting pages</h3></div>
      <div class="pbi-modal-body">
        <ul class="pbi-progress-list">
          ${pageNames.map((n, i) => `<li data-idx="${i}"><span class="pbi-progress-dot"></span>${escapeHtml(n)}</li>`).join('')}
        </ul>
      </div>
      <div class="pbi-modal-actions">
        <button class="pbi-modal-btn pbi-modal-btn-cancel">Cancel</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  let cancelled = false;
  overlay.querySelector('.pbi-modal-btn-cancel').addEventListener('click', () => {
    cancelled = true;
    overlay.remove();
    // Abort any in-flight capture wait, or the wizard's `await captureVisiblePage()`
    // never resolves and the sidebar stays hidden forever
    if (screenshotResolver) screenshotResolver(null);
    chrome.runtime.sendMessage({ action: 'cancelCapture' }, () => { void chrome.runtime.lastError; });
  });
  return {
    setStatus(idx, status) { // 'active' | 'done' | 'failed'
      const li = overlay.querySelector(`li[data-idx="${idx}"]`);
      if (li) li.className = status;
    },
    close() { overlay.remove(); },
    isCancelled() { return cancelled; },
  };
}

function waitForPageSettle(expectedKey, timeoutMs = 6000) {
  return new Promise((resolve) => {
    const start = Date.now();
    (function poll() {
      const settled = getPageKey() === expectedKey && getReportCanvas();
      if (settled) {
        setTimeout(resolve, 800); // PBI visuals animate in after the container exists
        return;
      }
      if (Date.now() - start > timeoutMs) { resolve(); return; }
      setTimeout(poll, 250);
    })();
  });
}

async function generateMultiPagePresentation(format) {
  const pages = getAnnotatedPages();
  if (pages.length === 0) return;
  const originalKey = getPageKey();
  const Nav = window.PowerBIAnnotatorPageNavigator;

  const sidebar = document.getElementById('pbi-annotator-sidebar');
  const toggleBtn = document.getElementById('pbi-toggle-btn');
  const sidebarWasOpen = sidebar.classList.contains('open');
  sidebar.classList.remove('open');

  const progress = showExportProgress(pages.map((p) => p.name));
  const pageDataList = [];
  let globalNumber = 1;

  for (let i = 0; i < pages.length; i++) {
    if (progress.isCancelled()) break;
    const page = pages[i];
    const pageAnnotations = allAnnotationsCache[page.key] || [];
    const sectionHash = page.key.split('#')[1] || null;
    progress.setStatus(i, 'active');

    if (page.key !== getPageKey()) {
      const navEl = Nav.findNavElement(document, {
        sectionHash,
        displayName: page.name,
      });
      if (!navEl) {
        progress.setStatus(i, 'failed');
        continue; // page listed as failed rather than silently blank
      }
      navEl.click();
      await waitForPageSettle(page.key);
    }

    // Hide our UI, capture, restore
    sidebar.style.display = 'none';
    toggleBtn.style.display = 'none';
    await new Promise((r) => setTimeout(r, 200));
    const raw = await captureVisiblePage();
    sidebar.style.display = '';
    toggleBtn.style.display = '';

    if (progress.isCancelled()) break;
    if (!raw) { progress.setStatus(i, 'failed'); continue; }
    const reportCanvas = getReportCanvas();
    const screenshot = reportCanvas ? await cropScreenshotToCanvas(raw, reportCanvas) : raw;

    pageDataList.push({
      pageName: page.name,
      screenshot,
      comments: pageAnnotations.map((a) => ({
        number: globalNumber++,
        comment: a.comment,
        date: new Date(a.timestamp).toLocaleDateString(),
        tool: a.tool || 'rectangle',
        color: a.color || '#0078d4',
      })),
    });
    progress.setStatus(i, 'done');
  }

  // Return to the page the user started on
  if (getPageKey() !== originalKey) {
    const homePage = pages.find((p) => p.key === originalKey);
    if (homePage) {
      const backEl = Nav.findNavElement(document, {
        sectionHash: homePage.key.split('#')[1] || null,
        displayName: homePage.name,
      });
      if (backEl) backEl.click();
    }
  }

  progress.close();
  if (sidebarWasOpen) sidebar.classList.add('open');
  if (pageDataList.length === 0) {
    await showModal('No pages could be captured. Click the extension icon when prompted and try again.');
    return;
  }
  if (format === 'ppt') await generateMultiPagePptx(pageDataList);
  else await generateMultiPagePdf(pageDataList);
  showToast('Export ready — check your downloads');
}

/**
 * Generate and download a real .pptx file using PptxGenJS.
 * Creates a widescreen slide with the screenshot and numbered comments side-by-side.
 */
async function generatePptx(screenshot, comments, pageName) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE'; // 13.33" x 7.5"

  const slideW = 13.33;
  const slideH = 7.5;
  const margin = 0.3;
  const contentW = slideW - margin * 2;

  // Title height
  const titleH = 0.4;
  const titleY = margin;

  // Content area after title
  const contentY = titleY + titleH + 0.2;
  const contentH = slideH - contentY - margin;

  // Split content area: 80% for screenshot, 20% for comments
  const screenshotW = contentW * 0.80;
  const commentsW = contentW * 0.18;
  const gap = contentW - screenshotW - commentsW; // spacing between screenshot and comments

  // First slide
  let slide = pres.addSlide();

  // Title
  slide.addText(pageName, {
    x: margin,
    y: titleY,
    w: contentW,
    h: titleH,
    fontSize: 18,
    bold: true,
    color: '0078d4',
    fontFace: 'Arial',
  });

  // Screenshot — left side, calculate dimensions to maintain aspect ratio
  if (screenshot) {
    const img = new Image();
    img.src = screenshot;
    const imgLoaded = await new Promise(resolve => {
      img.onload = () => resolve(true);
      img.onerror = () => resolve(false);
    });

    if (imgLoaded) {
      const fit = window.PowerBIAnnotatorPresentationLayout.computeImageFit(
        img.naturalWidth, img.naturalHeight, screenshotW, contentH);
      const imgW = fit.width;
      const imgH = fit.height;

      // Center the image vertically within the available area
      const imgX = margin;
      const imgY = contentY + (contentH - imgH) / 2;

      slide.addImage({
        data: screenshot,
        x: imgX,
        y: imgY,
        w: imgW,
        h: imgH,
      });
    } else {
      slide.addText('Screenshot capture failed', {
        x: margin,
        y: contentY,
        w: screenshotW,
        h: contentH,
        align: 'center',
        valign: 'middle',
        fontSize: 14,
        color: '999999',
        fontFace: 'Arial',
      });
    }
  } else {
    slide.addText('Screenshot capture failed', {
      x: margin,
      y: contentY,
      w: screenshotW,
      h: contentH,
      align: 'center',
      valign: 'middle',
      fontSize: 14,
      color: '999999',
      fontFace: 'Arial',
    });
  }

  // Comments section — right side
  const commentsX = margin + screenshotW + gap;
  
  // Add "Comments" header
  slide.addText('Comments', {
    x: commentsX,
    y: contentY,
    w: commentsW,
    h: 0.3,
    fontSize: 12,
    bold: true,
    color: '0078d4',
    fontFace: 'Arial',
  });

  // Comments list
  const commentsListY = contentY + 0.35;
  const commentsListH = contentH - 0.35;
  const commentLineH = 0.28;
  const maxCommentsPerSlide = window.PowerBIAnnotatorPresentationLayout.commentsPerSlide({
    availableHeight: commentsListH,
    lineHeight: commentLineH,
  });

  let currentY = commentsListY;
  let commentsOnSlide = 0;

  for (let i = 0; i < comments.length; i++) {
    // Check if we need a new slide for overflow
    if (commentsOnSlide >= maxCommentsPerSlide) {
      slide = pres.addSlide();
      slide.addText(pageName + ' (continued)', {
        x: margin,
        y: margin,
        w: contentW,
        h: titleH,
        fontSize: 18,
        bold: true,
        color: '0078d4',
        fontFace: 'Arial',
      });
      
      slide.addText('Comments (continued)', {
        x: margin,
        y: contentY,
        w: contentW,
        h: 0.3,
        fontSize: 12,
        bold: true,
        color: '0078d4',
        fontFace: 'Arial',
      });
      
      currentY = commentsListY;
      commentsOnSlide = 0;
    }

    const comment = comments[i];

    // Number badge + comment text
    slide.addText([
      { text: `${comment.number}  `, options: { bold: true, color: 'FFFFFF', fontSize: 9 } },
    ], {
      x: commentsX,
      y: currentY,
      w: 0.25,
      h: 0.25,
      fontFace: 'Arial',
      align: 'center',
      valign: 'middle',
      fill: { color: '0078d4' },
      shape: pres.ShapeType.ellipse,
    });

    // Comment text
    slide.addText(comment.comment, {
      x: commentsX + 0.28,
      y: currentY,
      w: commentsW - 0.28,
      h: commentLineH,
      fontSize: 8,
      color: '333333',
      fontFace: 'Arial',
      valign: 'top',
    });

    currentY += commentLineH;
    commentsOnSlide++;
  }

  // Download the .pptx file
  const now = new Date();
  const filename = `PowerBI_Pages_${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}.pptx`;

  await pres.writeFile({ fileName: filename });

  await showModal(`Exported page with ${comments.length} annotation${comments.length > 1 ? 's' : ''} to ${filename}\n\nThe .pptx file has been downloaded. You can open it directly in PowerPoint or Google Slides.`);
}

/**
 * Generate and download a real .pdf file using jsPDF.
 * Creates an A4 landscape page with the screenshot and numbered comments side-by-side.
 */
async function generatePdf(screenshot, comments, pageName) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4'
  });

  const pageW = 297; // A4 landscape width in mm
  const pageH = 210; // A4 landscape height in mm
  const margin = 10;
  const contentW = pageW - margin * 2;

  // Title section
  const titleH = 12;
  doc.setFontSize(18);
  doc.setTextColor(0, 120, 212); // #0078d4
  doc.text(pageName, margin, margin + 8);

  // Annotation count
  doc.setFontSize(10);
  doc.setTextColor(102, 102, 102); // #666
  const annotationText = `${comments.length} Annotation${comments.length > 1 ? 's' : ''}`;
  const textWidth = doc.getTextWidth(annotationText);
  doc.text(annotationText, pageW - margin - textWidth, margin + 8);

  // Content area
  const contentY = margin + titleH + 5;
  const contentH = pageH - contentY - margin;

  // Split content: 75% for screenshot, 25% for comments
  const screenshotW = contentW * 0.75;
  const commentsW = contentW * 0.25;
  const gap = 5;

  // Screenshot - left side
  if (screenshot) {
    const img = new Image();
    img.src = screenshot;
    const imgLoaded = await new Promise(resolve => {
      img.onload = () => resolve(true);
      img.onerror = () => resolve(false);
    });

    if (imgLoaded) {
      const fit = window.PowerBIAnnotatorPresentationLayout.computeImageFit(
        img.naturalWidth, img.naturalHeight, screenshotW, contentH);
      const imgW = fit.width;
      const imgH = fit.height;

      // Center vertically
      const imgX = margin;
      const imgY = contentY + (contentH - imgH) / 2;

      doc.addImage(screenshot, 'PNG', imgX, imgY, imgW, imgH);
    } else {
      doc.setFontSize(12);
      doc.setTextColor(153, 153, 153);
      doc.text('Screenshot capture failed', margin + screenshotW / 2, contentY + contentH / 2, { align: 'center' });
    }
  } else {
    // Placeholder for failed screenshot
    doc.setFontSize(12);
    doc.setTextColor(153, 153, 153);
    doc.text('Screenshot capture failed', margin + screenshotW / 2, contentY + contentH / 2, { align: 'center' });
  }

  // Comments section - right side
  const commentsX = margin + screenshotW + gap;

  // "Comments" header
  doc.setFontSize(14);
  doc.setTextColor(0, 120, 212);
  doc.text('Comments', commentsX, contentY + 5);

  // Comments list
  let currentY = contentY + 12;
  const lineHeight = 8;
  const badgeRadius = 3;

  for (let i = 0; i < comments.length; i++) {
    const comment = comments[i];

    // Check if we need a new page
    if (currentY + lineHeight > pageH - margin) {
      doc.addPage();
      currentY = margin + 10;

      // Add "Comments (continued)" header
      doc.setFontSize(14);
      doc.setTextColor(0, 120, 212);
      doc.text('Comments (continued)', margin, currentY);
      currentY += 10;
    }

    // Number badge (circle)
    doc.setFillColor(0, 120, 212); // #0078d4
    doc.circle(commentsX + badgeRadius, currentY - 1, badgeRadius, 'F');

    // Badge number
    doc.setFontSize(9);
    doc.setTextColor(255, 255, 255);
    doc.text(String(comment.number), commentsX + badgeRadius, currentY + 1, { align: 'center' });

    // Comment text
    doc.setFontSize(9);
    doc.setTextColor(51, 51, 51);
    const textX = commentsX + badgeRadius * 2 + 3;
    const maxWidth = commentsW - (badgeRadius * 2 + 3) - 2;
    
    // Split text into lines if too long
    const lines = doc.splitTextToSize(comment.comment, maxWidth);
    doc.text(lines, textX, currentY + 1);

    currentY += Math.max(lineHeight, lines.length * 4);
  }

  // Download the PDF
  const now = new Date();
  const filename = `PowerBI_Pages_${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}.pdf`;

  doc.save(filename);

  await showModal(`Exported page with ${comments.length} annotation${comments.length > 1 ? 's' : ''} to ${filename}\n\nThe PDF file has been downloaded automatically.`);
}

/**
 * Generate a multi-page PDF with all annotated pages.
 * Each page gets its own section: screenshot on left, comments on right.
 */
async function generateMultiPagePdf(pageDataList) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({
    orientation: 'landscape',
    unit: 'mm',
    format: 'a4'
  });

  const pageW = 297;
  const pageH = 210;
  const margin = 10;
  const contentW = pageW - margin * 2;
  const titleH = 12;

  let totalComments = 0;

  for (let p = 0; p < pageDataList.length; p++) {
    const pageData = pageDataList[p];
    totalComments += pageData.comments.length;

    // Add a new PDF page for each Power BI page (except the first)
    if (p > 0) {
      doc.addPage();
    }

    // Title section
    doc.setFontSize(18);
    doc.setTextColor(0, 120, 212);
    doc.text(pageData.pageName, margin, margin + 8);

    // Page indicator
    doc.setFontSize(10);
    doc.setTextColor(102, 102, 102);
    const pageLabel = `Page ${p + 1} of ${pageDataList.length} \u2022 ${pageData.comments.length} Annotation${pageData.comments.length !== 1 ? 's' : ''}`;
    const textWidth = doc.getTextWidth(pageLabel);
    doc.text(pageLabel, pageW - margin - textWidth, margin + 8);

    // Content area
    const contentY = margin + titleH + 5;
    const contentH = pageH - contentY - margin;
    const screenshotW = contentW * 0.75;
    const commentsW = contentW * 0.25;
    const gap = 5;

    // Screenshot - left side
    if (pageData.screenshot) {
      const img = new Image();
      img.src = pageData.screenshot;
      const imgLoaded = await new Promise(resolve => {
        img.onload = () => resolve(true);
        img.onerror = () => resolve(false);
      });

      if (imgLoaded) {
        const fit = window.PowerBIAnnotatorPresentationLayout.computeImageFit(
          img.naturalWidth, img.naturalHeight, screenshotW, contentH);
        const imgW = fit.width;
        const imgH = fit.height;
        const imgX = margin;
        const imgY = contentY + (contentH - imgH) / 2;
        doc.addImage(pageData.screenshot, 'PNG', imgX, imgY, imgW, imgH);
      } else {
        doc.setFontSize(12);
        doc.setTextColor(153, 153, 153);
        doc.text('Screenshot failed to load', margin + screenshotW / 2, contentY + contentH / 2, { align: 'center' });
      }
    } else {
      // No cached screenshot
      doc.setFontSize(12);
      doc.setTextColor(153, 153, 153);
      doc.text('No screenshot available', margin + screenshotW / 2, contentY + contentH / 2, { align: 'center' });
      doc.setFontSize(9);
      doc.text('(Visit and annotate this page to cache a screenshot)', margin + screenshotW / 2, contentY + contentH / 2 + 6, { align: 'center' });
    }

    // Comments section - right side
    const commentsX = margin + screenshotW + gap;
    doc.setFontSize(14);
    doc.setTextColor(0, 120, 212);
    doc.text('Comments', commentsX, contentY + 5);

    let currentY = contentY + 12;
    const lineHeight = 8;
    const badgeRadius = 3;

    for (let i = 0; i < pageData.comments.length; i++) {
      const comment = pageData.comments[i];

      // Check if we need a new PDF page for overflow
      if (currentY + lineHeight > pageH - margin) {
        doc.addPage();
        currentY = margin + 10;
        doc.setFontSize(14);
        doc.setTextColor(0, 120, 212);
        doc.text(pageData.pageName + ' \u2014 Comments (continued)', margin, currentY);
        currentY += 10;
      }

      // Number badge
      doc.setFillColor(0, 120, 212);
      doc.circle(commentsX + badgeRadius, currentY - 1, badgeRadius, 'F');
      doc.setFontSize(9);
      doc.setTextColor(255, 255, 255);
      doc.text(String(comment.number), commentsX + badgeRadius, currentY + 1, { align: 'center' });

      // Comment text
      doc.setFontSize(9);
      doc.setTextColor(51, 51, 51);
      const textX = commentsX + badgeRadius * 2 + 3;
      const maxWidth = commentsW - (badgeRadius * 2 + 3) - 2;
      const lines = doc.splitTextToSize(comment.comment, maxWidth);
      doc.text(lines, textX, currentY + 1);

      currentY += Math.max(lineHeight, lines.length * 4);
    }
  }

  const now = new Date();
  const filename = `PowerBI_AllPages_${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}.pdf`;
  doc.save(filename);

  await showModal(`Exported ${pageDataList.length} page(s) with ${totalComments} annotation(s) to ${filename}\n\nThe PDF file has been downloaded automatically.`);
}

/**
 * Generate a multi-page PPTX with all annotated pages.
 * Each page gets its own slide: screenshot on left, comments on right.
 */
async function generateMultiPagePptx(pageDataList) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';

  const slideW = 13.33;
  const slideH = 7.5;
  const margin = 0.3;
  const contentW = slideW - margin * 2;
  const titleH = 0.4;
  const titleY = margin;
  const contentY = titleY + titleH + 0.2;
  const contentH = slideH - contentY - margin;
  const screenshotW = contentW * 0.80;
  const commentsW = contentW * 0.18;
  const gap = contentW - screenshotW - commentsW;

  let totalComments = 0;

  for (let p = 0; p < pageDataList.length; p++) {
    const pageData = pageDataList[p];
    totalComments += pageData.comments.length;

    let slide = pres.addSlide();

    // Title with page indicator
    const titleText = pageData.pageName + (pageDataList.length > 1 ? ` (${p + 1}/${pageDataList.length})` : '');
    slide.addText(titleText, {
      x: margin, y: titleY, w: contentW, h: titleH,
      fontSize: 18, bold: true, color: '0078d4', fontFace: 'Arial',
    });

    // Screenshot
    if (pageData.screenshot) {
      const img = new Image();
      img.src = pageData.screenshot;
      const imgLoaded = await new Promise(resolve => {
        img.onload = () => resolve(true);
        img.onerror = () => resolve(false);
      });

      if (imgLoaded) {
        const fit = window.PowerBIAnnotatorPresentationLayout.computeImageFit(
          img.naturalWidth, img.naturalHeight, screenshotW, contentH);
        const imgW = fit.width;
        const imgH = fit.height;
        slide.addImage({
          data: pageData.screenshot,
          x: margin, y: contentY + (contentH - imgH) / 2,
          w: imgW, h: imgH,
        });
      } else {
        slide.addText('Screenshot failed to load', {
          x: margin, y: contentY, w: screenshotW, h: contentH,
          align: 'center', valign: 'middle', fontSize: 14, color: '999999', fontFace: 'Arial',
        });
      }
    } else {
      slide.addText('No screenshot available\n(Visit this page to cache a screenshot)', {
        x: margin, y: contentY, w: screenshotW, h: contentH,
        align: 'center', valign: 'middle', fontSize: 14, color: '999999', fontFace: 'Arial',
      });
    }

    // Comments section
    const commentsX = margin + screenshotW + gap;
    slide.addText('Comments', {
      x: commentsX, y: contentY, w: commentsW, h: 0.3,
      fontSize: 12, bold: true, color: '0078d4', fontFace: 'Arial',
    });

    const commentsListY = contentY + 0.35;
    const commentsListH = contentH - 0.35;
    const commentLineH = 0.28;
    const maxCommentsPerSlide = window.PowerBIAnnotatorPresentationLayout.commentsPerSlide({
    availableHeight: commentsListH,
    lineHeight: commentLineH,
  });
    let currentSlideY = commentsListY;
    let commentsOnSlide = 0;

    for (let i = 0; i < pageData.comments.length; i++) {
      if (commentsOnSlide >= maxCommentsPerSlide) {
        slide = pres.addSlide();
        slide.addText(pageData.pageName + ' (continued)', {
          x: margin, y: margin, w: contentW, h: titleH,
          fontSize: 18, bold: true, color: '0078d4', fontFace: 'Arial',
        });
        slide.addText('Comments (continued)', {
          x: margin, y: contentY, w: contentW, h: 0.3,
          fontSize: 12, bold: true, color: '0078d4', fontFace: 'Arial',
        });
        currentSlideY = commentsListY;
        commentsOnSlide = 0;
      }

      const comment = pageData.comments[i];

      // Number badge
      slide.addText([
        { text: `${comment.number}  `, options: { bold: true, color: 'FFFFFF', fontSize: 9 } },
      ], {
        x: commentsX, y: currentSlideY, w: 0.25, h: 0.25,
        fontFace: 'Arial', align: 'center', valign: 'middle',
        fill: { color: '0078d4' }, shape: pres.ShapeType.ellipse,
      });

      // Comment text
      slide.addText(comment.comment, {
        x: commentsX + 0.28, y: currentSlideY,
        w: commentsW - 0.28, h: commentLineH,
        fontSize: 8, color: '333333', fontFace: 'Arial', valign: 'top',
      });

      currentSlideY += commentLineH;
      commentsOnSlide++;
    }
  }

  const now = new Date();
  const filename = `PowerBI_AllPages_${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}.pptx`;
  await pres.writeFile({ fileName: filename });

  await showModal(`Exported ${pageDataList.length} page(s) with ${totalComments} annotation(s) to ${filename}\n\nThe .pptx file has been downloaded. You can open it directly in PowerPoint or Google Slides.`);
}

// Get the Power BI report canvas element
function getReportCanvas() {
  // Try multiple selectors to find the report canvas
  const canvasSelectors = [
    'div[class*="explorationContainer"]',
    'div.explorationContainer',
    'exploration-container',            // App view (workspace chrome absent)
    'report-embed div[class*="displayArea"]',
    'visual-container-repeat',
    'explore-canvas-modern',
    'explore-canvas',
    'iframe[title*="Report"]',
    '.reportCanvas',
    '.visualContainer'
  ];
  
  for (const selector of canvasSelectors) {
    const element = document.querySelector(selector);
    if (element) {
      return element;
    }
  }

  return null;
}

// Crop screenshot to only include the report canvas area
async function cropScreenshotToCanvas(screenshotDataUrl, canvasElement) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const rect = canvasElement.getBoundingClientRect();
      
      // Create a canvas to crop the image
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      
      // Account for device pixel ratio
      const dpr = window.devicePixelRatio || 1;
      
      // Set canvas size to the cropped area
      canvas.width = rect.width * dpr;
      canvas.height = rect.height * dpr;

      // Calculate source coordinates using the actual image-to-viewport ratio
      // captureVisibleTab returns an image sized to viewport * dpr, so we map
      // viewport coordinates directly to image pixels.
      const scaleX = img.width / window.innerWidth;
      const scaleY = img.height / window.innerHeight;

      const sourceX = rect.left * scaleX;
      const sourceY = rect.top * scaleY;
      const sourceWidth = rect.width * scaleX;
      const sourceHeight = rect.height * scaleY;
      
      // Draw the cropped portion
      ctx.drawImage(
        img,
        sourceX, sourceY, sourceWidth, sourceHeight,
        0, 0, canvas.width, canvas.height
      );
      
      // Convert back to data URL
      const croppedDataUrl = canvas.toDataURL('image/png');
      resolve(croppedDataUrl);
    };
    
    img.onerror = () => {
      console.error('Failed to load screenshot for cropping');
      resolve(screenshotDataUrl); // Return original if cropping fails
    };
    
    img.src = screenshotDataUrl;
  });
}

/**
 * Extract a clean display name from a DOM element.
 * Handles aria-label, title, and textContent with Power BI suffix cleanup.
 */
function extractPageNameFromElement(element) {
  let pageName = element.getAttribute('aria-label') ||
                 element.getAttribute('title') ||
                 element.textContent?.trim();

  if (pageName && pageName !== 'Page navigation' && !pageName.toLowerCase().includes('page navigation')) {
    pageName = pageName
      .replace(/[,\s]+selected$/i, '')
      .replace(/[,\s]+active$/i, '')
      .replace(/[,\s]+current$/i, '')
      .trim();

    if (pageName && pageName.length > 0) {
      return pageName;
    }
  }
  return null;
}

function getPageName() {
  // Priority 0: Use Power BI embed API cached page name (most reliable)
  const currentUrl = window.location.pathname + window.location.search;
  if (pageNameCache[currentUrl]) {
    return pageNameCache[currentUrl];
  }

  // Strategy 0: For Power BI Apps, use the left-nav tree directly.
  // This must come first — app nav is structurally different from workspace
  // report bottom tabs, and the broad selectors below can return the wrong
  // element (e.g. the report group header instead of the active page).
  if (window.location.pathname.includes('/apps/')) {
    const appPageName = getActiveAppNavPageName();
    if (appPageName) return appPageName;
  }

  // Strategy 1: Scoped search inside known page navigation containers.
  // This avoids false matches from other buttons on the page (e.g. filter pane,
  // visual action buttons) that may also carry aria-selected="true".
  // Power BI uses different container class names across workspace reports vs apps.
  const navContainerSelectors = [
    '.pagesNavigation',
    '.pagesNav',
    '.pages-navigation',
    '[aria-label="Page navigation"]',
    '[class*="pagesNavigation"]',
    '[class*="pageNavigation"]',
    '[class*="pagesNav"]',
    'nav[class*="page"]',
  ];

  const activeAttrSelectors = [
    'button[aria-selected="true"]',
    'button[aria-current="true"]',
    'button[aria-current="page"]',
    'button.active',
    'button.selected',
    'button[class*="isSelected"]',
    'button[class*="is-selected"]',
  ];

  for (const containerSel of navContainerSelectors) {
    const container = document.querySelector(containerSel);
    if (!container) continue;

    for (const activeSel of activeAttrSelectors) {
      const element = container.querySelector(activeSel);
      if (element) {
        const name = extractPageNameFromElement(element);
        if (name) return name;
      }
    }

    // Also try non-button active items (span/div with active class) inside nav
    const activeItem = container.querySelector(
      '.active, .selected, [class*="isActive"], [class*="is-active"]'
    );
    if (activeItem) {
      const name = extractPageNameFromElement(activeItem);
      if (name) return name;
    }
  }

  // Strategy 2: Broader unscoped selectors as fallback.
  // Order matters — more specific first to reduce false positives.
  const broadSelectors = [
    // Power BI app-specific patterns
    '[class*="pageItem"][class*="active"]',
    '[class*="pageItem"][class*="selected"]',
    '[data-page-index][aria-selected="true"]',
    // Standard ARIA patterns
    'button[role="tab"][aria-selected="true"][aria-label]',
    'button[role="tab"][aria-current="true"][aria-label]',
    'button[role="tab"][aria-current="page"][aria-label]',
    // Legacy Power BI selectors
    '.navigationPane .active button[title]',
    '.navigationPane .selected button[title]',
    '.navigationPane button.is-selected[title]',
    '.pagesNav .active .itemName',
    '.navigationPane .itemContainer.active .itemName',
    // Generic fallbacks
    'button[aria-selected="true"][aria-label]',
    'button[aria-current="true"][aria-label]',
    'div[role="tablist"] button[aria-selected="true"]',
  ];

  for (const selector of broadSelectors) {
    const element = document.querySelector(selector);
    if (element) {
      const name = extractPageNameFromElement(element);
      if (name) return name;
    }
  }

  // Strategy 3: URL query param — some embed URLs carry pageName explicitly.
  const searchParams = new URLSearchParams(window.location.search);
  const urlPageName = searchParams.get('pageName');
  if (urlPageName) {
    // pageName in URLs is usually "ReportSection{hash}" (internal ID), not display name.
    // Try to find a DOM element tagged with this value before giving up.
    const matchEl = document.querySelector(
      `[data-page-name="${urlPageName}"], [data-reportpage="${urlPageName}"]`
    );
    if (matchEl) {
      const name = extractPageNameFromElement(matchEl);
      if (name) return name;
    }
    // Don't return the raw ReportSection hash — it's not human-readable.
  }

  // Strategy 4: Document title — strip " - Power BI" and report name suffix.
  if (document.title) {
    const cleanTitle = document.title
      .replace(/ - Power BI.*$/i, '')
      .replace(/ \| Power BI.*$/i, '')
      .trim();
    if (cleanTitle && cleanTitle.length > 5) {
      return cleanTitle;
    }
  }

  // Strategy 5: Last path segment from URL.
  const path = window.location.pathname;
  const parts = path.split('/').filter(p => p);
  if (parts.length > 0) {
    return parts[parts.length - 1];
  }

  return 'Power BI Report';
}

// Create annotation element from annotation data
// [Fix #1] Uses appendChild instead of innerHTML += to preserve SVG namespace
// [Fix #2] Uses startPoint/endPoint for correct arrow/line direction
// Resolve an annotation to pixel space for the current layout.
// v2 annotations rescale to wherever the canvas is now; v1 annotations
// (or no canvas found) render at their stored absolute position.
function resolveAnnotationForLayout(annotation) {
  const Coords = window.PowerBIAnnotatorCoords;
  const canvas = getReportCanvas();
  if (canvas && annotation.coordSpace === 'canvas' && annotation.rel) {
    return Coords.annotationToAbsolute(annotation, Coords.getCanvasPageRect(canvas, window));
  }
  return annotation;
}

function renderAnnotationsForCurrentPage() {
  document.querySelectorAll('.pbi-annotation-box').forEach((box) => box.remove());
  const globalStart = getGlobalStartNumber();
  annotations.forEach((annotation, index) => {
    const box = createAnnotationElement(annotation, globalStart + index + 1);
    document.body.appendChild(box);
  });
}

// One-time upgrade of v1 (absolute-pixel) annotations. Uses the current
// canvas rect: correct whenever the layout still matches draw-time, and
// no worse than the old behavior when it doesn't.
function migrateLoadedAnnotations() {
  const Coords = window.PowerBIAnnotatorCoords;
  const canvas = getReportCanvas();
  if (!canvas) return false;
  const rect = Coords.getCanvasPageRect(canvas, window);
  let changed = false;
  annotations.forEach((a, i) => {
    const migrated = Coords.migrateAnnotation(a, rect);
    if (migrated !== a) {
      annotations[i] = migrated;
      changed = true;
    }
  });
  return changed;
}

function createAnnotationElement(annotation, number) {
  const resolved = resolveAnnotationForLayout(annotation);
  const box = document.createElement("div");
  box.className = "pbi-annotation-box";
  box.dataset.id = annotation.id;
  box.style.left = resolved.x + "px";
  box.style.top = resolved.y + "px";
  box.style.width = resolved.width + "px";
  box.style.height = resolved.height + "px";

  const toolName = annotation.tool || 'rectangle';
  const color = annotation.color || '#0078d4';
  box.style.borderColor = color;

  const Tools = window.PowerBIAnnotatorTools;
  const tool = Tools[toolName];
  if (tool) {
    const geometry = Tools.geometryFromAnnotation(resolved);
    const svg = tool.render(geometry, color);
    if (svg) box.appendChild(svg);
  }

  const badge = document.createElement('div');
  badge.className = 'pbi-annotation-number';
  badge.textContent = number;
  badge.style.background = color;
  box.appendChild(badge);

  box.addEventListener("click", (e) => {
    e.stopPropagation();
    showAnnotationComment(annotation.id);
  });

  return box;
}

// Save annotations to Chrome storage [Fix #4] - uses cache to avoid read-modify-write race
function saveAnnotations() {
  // Delegates to PageStore. The pageKey, in-memory cache, and chrome.storage
  // write all live behind that seam.
  if (pageStore) {
    pageStore.saveAnnotations(annotations);
  }
}

function loadAnnotations() {
  pageStore.init().then(() => {
    allAnnotationsCache = pageStore._snapshot();
    annotations = pageStore.current().annotations;

    // Migrate old annotations to include pageName field
    let needsSave = false;
    const currentPageName = getPageName();
    annotations.forEach(annotation => {
      if (!annotation.pageName) {
        annotation.pageName = currentPageName;
        needsSave = true;
      }
    });
    if (needsSave) {
      pageStore.saveAnnotations(annotations);
    }

    if (migrateLoadedAnnotations()) {
      pageStore.saveAnnotations(annotations);
    }
    renderAnnotationsForCurrentPage();

    renderComments();
    renderPageList();
  });
}

// Extract report ID from Power BI URL
function getReportId() {
  // Power BI URL format: /groups/{workspace}/reports/{reportId}/ReportSection...
  // or: /reportEmbed?reportId={reportId}
  const pathname = window.location.pathname;
  const search = window.location.search;
  
  // Try to extract from path
  const pathMatch = pathname.match(/\/reports\/([^\/]+)/);
  if (pathMatch) {
    return pathMatch[1];
  }
  
  // Try to extract from query params
  const searchParams = new URLSearchParams(search);
  const reportId = searchParams.get('reportId');
  if (reportId) {
    return reportId;
  }
  
  // Fallback: use full pathname as report ID
  return pathname.split('/').filter(p => p).join('_');
}

// Get unique key for current page [Fix #5] - includes query params for Power BI page navigation
// Returns the canonical key used by PageStore so legacy callers and PageStore
// agree on what identifies "the current page". Falls back to pathname+search
// before PageStore is initialised (very brief window at startup).
function getPageKey() {
  if (pageStore) return pageStore.current().key;
  return window.location.pathname + window.location.search;
}

// Get the active page name from the Power BI App left-nav sidebar.
// 
// Strategy A (most reliable): match the current URL's ReportSection hash
// against nav link hrefs. The sidebar <a> tags point to the same
// ReportSection{hash} URLs, so we can find the exact nav item by href
// and read its display text — no guessing CSS classes needed.
//
// Strategy B (fallback): ARIA / class-based selectors for cases where
// nav links aren't anchor tags or hrefs aren't available.
function getActiveAppNavPageName() {
  // Strategy A: URL hash → nav link text match
  const sectionMatch = window.location.pathname.match(/\/(ReportSection[a-zA-Z0-9]+)/);
  if (sectionMatch) {
    const sectionId = sectionMatch[1];
    // Find all anchor tags whose href contains this ReportSection ID
    const links = document.querySelectorAll(`a[href*="${sectionId}"]`);
    for (const link of links) {
      const text = link.textContent?.trim();
      if (text && text.length > 0 && !text.includes('ReportSection')) {
        return text;
      }
    }
  }

  // Strategy B: ARIA / class selectors for tree nav
  const selectors = [
    '[role="treeitem"][aria-selected="true"]',
    '[role="treeitem"][aria-current="page"]',
    '[role="treeitem"][aria-current="true"]',
    '[class*="navItem"][class*="isSelected"]',
    '[class*="navItem"][class*="is-selected"]',
    '[class*="navItem"][class*="active"]',
    '[class*="navItem"][class*="selected"]',
    'li[class*="active"] a',
    'li[class*="selected"] a',
    'li[class*="isSelected"] a',
  ];

  for (const sel of selectors) {
    const el = document.querySelector(sel);
    if (el) {
      const name = extractPageNameFromElement(el);
      if (name) return name;
    }
  }
  return null;
}

// Format timestamp
function formatTime(timestamp) {
  const date = new Date(timestamp);
  const now = new Date();
  const diffMinutes = Math.floor((now - date) / 60000);

  if (diffMinutes < 1) return "Just now";
  if (diffMinutes < 60) return `${diffMinutes}m ago`;
  if (diffMinutes < 1440) return `${Math.floor(diffMinutes / 60)}h ago`;
  return date.toLocaleDateString();
}

// Escape HTML to prevent XSS
function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

// Initialize when DOM is ready
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", init);
} else {
  init();
}
