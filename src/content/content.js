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
let allAnnotationsCache = null; // [Fix #4] Cache to avoid read-modify-write race
let annotationIdCounter = 0; // [Fix #11] Counter to avoid Date.now() collisions
let lastPageKey = null; // Track current page key for SPA navigation detection
let screenshotCache = {}; // { pageKey: dataUrl } - cached screenshots per page

// --- Custom Modal Helpers (Fix #8: replace blocking prompt/alert/confirm) ---

/**
 * Show an informational modal (replaces alert).
 * Returns a promise that resolves when the user clicks OK.
 */
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
 * Update the number badges on all annotation boxes to match current array order. [Fix #3]
 * Called after deletion to keep page badges in sync with the sidebar.
 */
function renumberAnnotations() {
  annotations.forEach((annotation, index) => {
    const box = document.querySelector(`.pbi-annotation-box[data-id="${annotation.id}"]`);
    if (box) {
      const badge = box.querySelector('.pbi-annotation-number');
      if (badge) {
        badge.textContent = index + 1;
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

// Initialize the extension
function init() {
  createSidebar();
  loadAnnotations();
  loadScreenshotCache();
  setupEventListeners();
  lastPageKey = getPageKey();
  startNavigationWatcher();
  console.log("Power BI Annotator initialized");
}

// Poll for URL changes to detect SPA navigation (Power BI changes URL without page reload)
function startNavigationWatcher() {
  setInterval(() => {
    const currentKey = getPageKey();
    if (currentKey !== lastPageKey) {
      onPageChanged(lastPageKey, currentKey);
    }
  }, 500);
}

// Handle SPA page navigation: save current state, clear DOM, load new page's annotations
function onPageChanged(oldKey, newKey) {
  // Save current page's annotations under the old key
  saveAnnotations();

  // Cache a screenshot of the page we're leaving (best-effort, may already be transitioning)
  cacheCurrentScreenshot(oldKey);

  // Turn off annotation mode if active
  if (isAnnotationMode) {
    toggleAnnotationMode();
  }

  // Clear annotation DOM elements from old page
  document.querySelectorAll('.pbi-annotation-box').forEach(box => box.remove());

  // Load new page's annotations from in-memory cache
  if (allAnnotationsCache) {
    annotations = allAnnotationsCache[newKey] || [];
  } else {
    annotations = [];
  }

  // Render new page's annotations
  annotations.forEach((annotation, index) => {
    const box = createAnnotationElement(annotation, index + 1);
    document.body.appendChild(box);
  });

  lastPageKey = newKey;
  renderComments();
  renderPageList();
}

// Request a silent screenshot from the background script (uses host_permissions, no user gesture needed)
function cacheCurrentScreenshot(pageKey) {
  const key = pageKey || getPageKey();
  // Don't cache if page has no annotations
  if (allAnnotationsCache && (!allAnnotationsCache[key] || allAnnotationsCache[key].length === 0)) {
    return;
  }
  try {
    chrome.runtime.sendMessage({ action: 'captureForCache' }, (response) => {
      if (chrome.runtime.lastError) {
        console.log('Screenshot cache skipped:', chrome.runtime.lastError.message);
        return;
      }
      if (response && response.screenshot) {
        screenshotCache[key] = response.screenshot;
        // Limit to 20 most recent pages to manage storage size
        const keys = Object.keys(screenshotCache);
        if (keys.length > 20) {
          delete screenshotCache[keys[0]];
        }
        chrome.storage.local.set({ screenshotCache });
      }
    });
  } catch (e) {
    console.log('Screenshot cache error:', e);
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
      <h3>Comments</h3>
      <button id="pbi-close-sidebar" class="pbi-btn-close">\u00d7</button>
    </div>
    <div class="pbi-sidebar-controls">
      <button id="pbi-toggle-annotate" class="pbi-btn pbi-btn-full">
        \ud83d\udccd Start Annotating
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
        <input type="color" id="pbi-color-picker" value="#0078d4" title="Color">
      </div>
      <div class="pbi-button-row">
        <button id="pbi-export-pages" class="pbi-btn pbi-btn-primary">
          \ud83d\udcf8 Export Pages
        </button>
        <button id="pbi-export-annotations" class="pbi-btn pbi-btn-success">
          \ud83d\udcca Export CSV
        </button>
        <button id="pbi-clear-all" class="pbi-btn pbi-btn-danger">
          Clear
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
      <p class="pbi-empty-state">No comments yet. Click "Start Annotating" to begin.</p>
    </div>
  `;

  document.body.appendChild(sidebar);

  // Create toggle button
  const toggleBtn = document.createElement("button");
  toggleBtn.id = "pbi-toggle-btn";
  toggleBtn.className = "pbi-toggle-btn";
  toggleBtn.innerHTML = "\ud83d\udcac";
  toggleBtn.title = "Toggle Comments Sidebar";
  document.body.appendChild(toggleBtn);
}

// Setup event listeners
function setupEventListeners() {
  // Toggle sidebar
  document
    .getElementById("pbi-toggle-btn")
    .addEventListener("click", toggleSidebar);
  document
    .getElementById("pbi-close-sidebar")
    .addEventListener("click", toggleSidebar);

  // Toggle annotation mode
  document
    .getElementById("pbi-toggle-annotate")
    .addEventListener("click", toggleAnnotationMode);

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

  // Color picker
  document.getElementById("pbi-color-picker").addEventListener("change", (e) => {
    currentColor = e.target.value;
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
  const btn = document.getElementById("pbi-toggle-annotate");
  const toolbar = document.getElementById("pbi-drawing-toolbar");

  if (isAnnotationMode) {
    btn.textContent = "\u2713 Annotating (Click & Drag)";
    btn.classList.add("active");
    toolbar.style.display = "flex";
    document.body.style.cursor = "crosshair";
  } else {
    btn.textContent = "\ud83d\udccd Start Annotating";
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
    currentAnnotation.innerHTML = `<svg class="pbi-freehand-svg" style="position: absolute; top: 0; left: 0; pointer-events: none;"><path d="" stroke="${currentColor}" stroke-width="3" fill="none"/></svg>`;
  } else {
    currentAnnotation.style.width = "0px";
    currentAnnotation.style.height = "0px";
  }

  document.body.appendChild(currentAnnotation);
}

// Handle mouse move - resize annotation
function handleMouseMove(e) {
  if (!currentAnnotation) return;

  const width = e.pageX - startX;
  const height = e.pageY - startY;

  if (currentDrawingTool === 'freehand') {
    // Add point to freehand path
    freehandPoints.push({ x: e.pageX, y: e.pageY });

    // Update SVG path
    const svg = currentAnnotation.querySelector('svg');
    const path = svg.querySelector('path');

    // Calculate bounds
    const minX = Math.min(...freehandPoints.map(p => p.x));
    const minY = Math.min(...freehandPoints.map(p => p.y));
    const maxX = Math.max(...freehandPoints.map(p => p.x));
    const maxY = Math.max(...freehandPoints.map(p => p.y));

    // Update container position and size
    currentAnnotation.style.left = minX + "px";
    currentAnnotation.style.top = minY + "px";
    currentAnnotation.style.width = (maxX - minX) + "px";
    currentAnnotation.style.height = (maxY - minY) + "px";

    // Create path data (relative to container)
    const pathData = freehandPoints.map((p, i) => {
      const x = p.x - minX;
      const y = p.y - minY;
      return i === 0 ? `M ${x} ${y}` : `L ${x} ${y}`;
    }).join(' ');

    svg.setAttribute('width', maxX - minX);
    svg.setAttribute('height', maxY - minY);
    path.setAttribute('d', pathData);
  } else {
    // For other tools, draw bounding box
    currentAnnotation.style.width = Math.abs(width) + "px";
    currentAnnotation.style.height = Math.abs(height) + "px";
    currentAnnotation.style.left = (width < 0 ? e.pageX : startX) + "px";
    currentAnnotation.style.top = (height < 0 ? e.pageY : startY) + "px";

    // Update visual representation based on tool
    updateAnnotationVisual(currentAnnotation, width, height);
  }
}

// Update annotation visual based on drawing tool
function updateAnnotationVisual(element, width, height) {
  const absWidth = Math.abs(width);
  const absHeight = Math.abs(height);

  // Remove existing SVG if any
  const existingSvg = element.querySelector('svg');
  if (existingSvg) existingSvg.remove();

  if (currentDrawingTool === 'arrow') {
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute('width', absWidth);
    svg.setAttribute('height', absHeight);
    svg.style.position = 'absolute';
    svg.style.top = '0';
    svg.style.left = '0';
    svg.style.pointerEvents = 'none';

    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    const startX = width < 0 ? absWidth : 0;
    const startY = height < 0 ? absHeight : 0;
    const endX = width < 0 ? 0 : absWidth;
    const endY = height < 0 ? 0 : absHeight;

    // Arrow head size
    const headLen = 20;
    const angle = Math.atan2(endY - startY, endX - startX);
    const arrowX1 = endX - headLen * Math.cos(angle - Math.PI / 6);
    const arrowY1 = endY - headLen * Math.sin(angle - Math.PI / 6);
    const arrowX2 = endX - headLen * Math.cos(angle + Math.PI / 6);
    const arrowY2 = endY - headLen * Math.sin(angle + Math.PI / 6);

    path.setAttribute('d', `M ${startX} ${startY} L ${endX} ${endY} M ${arrowX1} ${arrowY1} L ${endX} ${endY} L ${arrowX2} ${arrowY2}`);
    path.setAttribute('stroke', currentColor);
    path.setAttribute('stroke-width', '3');
    path.setAttribute('fill', 'none');

    svg.appendChild(path);
    element.appendChild(svg);
  } else if (currentDrawingTool === 'line') {
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute('width', absWidth);
    svg.setAttribute('height', absHeight);
    svg.style.position = 'absolute';
    svg.style.top = '0';
    svg.style.left = '0';
    svg.style.pointerEvents = 'none';

    const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
    line.setAttribute('x1', width < 0 ? absWidth : 0);
    line.setAttribute('y1', height < 0 ? absHeight : 0);
    line.setAttribute('x2', width < 0 ? 0 : absWidth);
    line.setAttribute('y2', height < 0 ? 0 : absHeight);
    line.setAttribute('stroke', currentColor);
    line.setAttribute('stroke-width', '3');

    svg.appendChild(line);
    element.appendChild(svg);
  } else if (currentDrawingTool === 'circle') {
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute('width', absWidth);
    svg.setAttribute('height', absHeight);
    svg.style.position = 'absolute';
    svg.style.top = '0';
    svg.style.left = '0';
    svg.style.pointerEvents = 'none';

    const ellipse = document.createElementNS("http://www.w3.org/2000/svg", "ellipse");
    ellipse.setAttribute('cx', absWidth / 2);
    ellipse.setAttribute('cy', absHeight / 2);
    ellipse.setAttribute('rx', absWidth / 2);
    ellipse.setAttribute('ry', absHeight / 2);
    ellipse.setAttribute('stroke', currentColor);
    ellipse.setAttribute('stroke-width', '3');
    ellipse.setAttribute('fill', 'none');

    svg.appendChild(ellipse);
    element.appendChild(svg);
  }
  // Rectangle is handled by default border
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
    const annotation = {
      id: generateAnnotationId(),
      x: parseInt(finishedAnnotation.style.left),
      y: parseInt(finishedAnnotation.style.top),
      width: rect.width,
      height: rect.height,
      comment: comment.trim(),
      timestamp: new Date().toISOString(),
      url: window.location.href,
      tool: toolUsed,
      color: colorUsed,
      freehandPath: capturedFreehandPoints,
      startPoint: { x: drawStartX, y: drawStartY },
      endPoint: { x: drawEndX, y: drawEndY },
    };

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
    badge.textContent = annotations.length;
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
    showModal(
      `Comment #${annotations.indexOf(annotation) + 1}:\n\n${annotation.comment}`,
    );
  }
}

// Render comments in sidebar
function renderComments() {
  const commentsList = document.getElementById("pbi-comments-list");

  if (annotations.length === 0) {
    commentsList.innerHTML =
      '<p class="pbi-empty-state">No comments yet. Click "Start Annotating" to begin.</p>';
    return;
  }

  commentsList.innerHTML = annotations
    .map(
      (annotation, index) => `
    <div class="pbi-comment-item" data-id="${annotation.id}">
      <div class="pbi-comment-header">
        <span class="pbi-comment-number">#${index + 1}</span>
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

// Get list of all pages that have annotations
function getAnnotatedPages() {
  if (!allAnnotationsCache) return [];
  const currentKey = getPageKey();
  return Object.keys(allAnnotationsCache)
    .filter(key => allAnnotationsCache[key].length > 0)
    .map(key => {
      const pageAnnotations = allAnnotationsCache[key];
      // Derive page name from the first annotation's stored URL, or from the key
      let name = key;
      if (pageAnnotations.length > 0 && pageAnnotations[0].url) {
        try {
          const url = new URL(pageAnnotations[0].url);
          // Try document title pattern first (for current page)
          if (key === currentKey) {
            name = getPageName();
          } else {
            // Extract a readable name from the URL path
            const parts = url.pathname.split('/').filter(p => p);
            name = parts.length > 0 ? decodeURIComponent(parts[parts.length - 1]) : key;
          }
        } catch (e) {
          name = key;
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

// Clear all annotations [Fix #8] - async for custom confirm
async function clearAllAnnotations() {
  const confirmed = await showConfirm("Delete all comments? This cannot be undone.");
  if (!confirmed) return;

  annotations = [];
  document
    .querySelectorAll(".pbi-annotation-box")
    .forEach((box) => box.remove());
  saveAnnotations();
  renderComments();
  renderPageList();
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

// Build CSV rows for a set of pages (used by both single and multi-page export)
function buildCsvRows(pages) {
  const headers = ["No", "Page Name", "Date", "Comment"];
  const allRows = [];
  let globalNumber = 1;

  for (const page of pages) {
    for (const annotation of page.annotations) {
      const date = new Date(annotation.timestamp);
      allRows.push([
        globalNumber++,
        page.name,
        date.toLocaleDateString(),
        annotation.comment,
      ]);
    }
  }

  return [
    headers.join(","),
    ...allRows.map((row) =>
      row
        .map((cell) => {
          const cellStr = String(cell);
          if (cellStr.includes(",") || cellStr.includes('"') || cellStr.includes("\n")) {
            return `"${cellStr.replace(/"/g, '""')}"`;
          }
          return cellStr;
        })
        .join(","),
    ),
  ].join("\n");
}

// Export annotations to Excel (CSV format) [Fix #8, #10]
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
      annotations: allAnnotationsCache[p.key] || []
    }));
  } else {
    pages = [{ name: getPageName(), annotations }];
  }

  const totalCount = pages.reduce((sum, p) => sum + p.annotations.length, 0);
  const csvContent = buildCsvRows(pages);

  // Create blob and download
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);

  const now = new Date();
  const suffix = scope === 'all' ? 'AllPages' : 'Comments';
  const filename = `PowerBI_${suffix}_${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}-${String(now.getMinutes()).padStart(2, "0")}.csv`;

  link.setAttribute("href", url);
  link.setAttribute("download", filename);
  link.style.visibility = "hidden";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url); // [Fix #10] Prevent memory leak

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
    return;
  }

  // PDF format: generate real .pdf file using jsPDF
  await generatePdf(screenshot, comments, pageName);
}

/**
 * Generate a multi-page presentation/PDF that includes all annotated pages.
 * Uses cached screenshots for other pages and captures a fresh screenshot for the current page.
 */
async function generateMultiPagePresentation(format) {
  const pages = getAnnotatedPages();
  const currentKey = getPageKey();

  // Capture fresh screenshot for current page
  const sidebar = document.getElementById('pbi-annotator-sidebar');
  const toggleBtn = document.getElementById('pbi-toggle-btn');
  const sidebarWasOpen = sidebar.classList.contains('open');

  sidebar.style.display = 'none';
  toggleBtn.style.display = 'none';

  document.querySelectorAll('.pbi-annotation-box').forEach(box => {
    box.style.opacity = '1';
    box.style.zIndex = '';
  });

  const reportCanvas = getReportCanvas();

  await new Promise(resolve => setTimeout(resolve, 300));

  // Try to capture current page screenshot
  let currentScreenshot = null;
  try {
    chrome.runtime.sendMessage({ action: 'prepareCapture' }, (response) => {
      if (chrome.runtime.lastError) {
        console.error('Could not prepare capture:', chrome.runtime.lastError.message);
      }
    });

    const result = await waitForScreenshotOrCancel();
    if (result === null) {
      sidebar.style.display = '';
      toggleBtn.style.display = '';
      if (sidebarWasOpen) sidebar.classList.add('open');
      return;
    }
    if (result.screenshot && reportCanvas) {
      currentScreenshot = await cropScreenshotToCanvas(result.screenshot, reportCanvas);
    } else if (result.screenshot) {
      currentScreenshot = result.screenshot;
    }
  } catch (error) {
    console.error('Failed to capture screenshot:', error);
  }

  // Restore sidebar
  sidebar.style.display = '';
  toggleBtn.style.display = '';
  if (sidebarWasOpen) sidebar.classList.add('open');

  // Update screenshot cache for current page
  if (currentScreenshot) {
    screenshotCache[currentKey] = currentScreenshot;
    chrome.storage.local.set({ screenshotCache });
  }

  // Build page data array with screenshots and comments
  let globalNumber = 1;
  const pageDataList = pages.map(page => {
    const pageAnnotations = allAnnotationsCache[page.key] || [];
    const comments = pageAnnotations.map(annotation => ({
      number: globalNumber++,
      comment: annotation.comment,
      date: new Date(annotation.timestamp).toLocaleDateString(),
      tool: annotation.tool || 'rectangle',
      color: annotation.color || '#0078d4'
    }));
    return {
      pageName: page.name,
      screenshot: screenshotCache[page.key] || null,
      comments
    };
  });

  if (format === 'ppt') {
    await generateMultiPagePptx(pageDataList);
  } else {
    await generateMultiPagePdf(pageDataList);
  }
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
      const imgAspect = img.naturalWidth / img.naturalHeight;
      const boxAspect = screenshotW / contentH;

      let imgW, imgH;
      if (imgAspect > boxAspect) {
        // Image is wider than box — fit to width
        imgW = screenshotW;
        imgH = screenshotW / imgAspect;
      } else {
        // Image is taller than box — fit to height
        imgH = contentH;
        imgW = contentH * imgAspect;
      }

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
  const maxCommentsPerSlide = Math.floor(commentsListH / commentLineH);

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
      const imgAspect = img.naturalWidth / img.naturalHeight;
      const boxAspect = screenshotW / contentH;

      let imgW, imgH;
      if (imgAspect > boxAspect) {
        // Image is wider - fit to width
        imgW = screenshotW;
        imgH = screenshotW / imgAspect;
      } else {
        // Image is taller - fit to height
        imgH = contentH;
        imgW = contentH * imgAspect;
      }

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
        const imgAspect = img.naturalWidth / img.naturalHeight;
        const boxAspect = screenshotW / contentH;
        let imgW, imgH;
        if (imgAspect > boxAspect) {
          imgW = screenshotW;
          imgH = screenshotW / imgAspect;
        } else {
          imgH = contentH;
          imgW = contentH * imgAspect;
        }
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
        const imgAspect = img.naturalWidth / img.naturalHeight;
        const boxAspect = screenshotW / contentH;
        let imgW, imgH;
        if (imgAspect > boxAspect) {
          imgW = screenshotW;
          imgH = screenshotW / imgAspect;
        } else {
          imgH = contentH;
          imgW = contentH * imgAspect;
        }
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
    const maxCommentsPerSlide = Math.floor(commentsListH / commentLineH);
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
    'visual-container-repeat',
    'explore-canvas-modern',
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

function getPageName() {
  // Try to get the active page name from Power BI's Pages panel
  // Look for the selected/active page in the navigation pane
  const activePageSelectors = [
    'button[aria-selected="true"][aria-label]',
    'button[aria-selected="true"][title]',
    '.navigationPane .active button[title]',
    '.navigationPane .selected button[title]',
    '.navigationPane button.is-selected[title]',
    '.pagesNav .active .itemName',
    'button[role="tab"][aria-selected="true"]',
    '.navigationPane .itemContainer.active .itemName',
    'div[role="tablist"] button[aria-selected="true"]'
  ];
  
  for (const selector of activePageSelectors) {
    const element = document.querySelector(selector);
    if (element) {
      // Try aria-label first (works when navigation is collapsed)
      let pageName = element.getAttribute('aria-label') || 
                      element.getAttribute('title') || 
                      element.textContent?.trim();
      
      if (pageName && pageName !== 'Page navigation') {
        // Clean up common suffixes from Power BI
        pageName = pageName
          .replace(/[,\s]+selected$/i, '')
          .replace(/[,\s]+active$/i, '')
          .trim();
        
        if (pageName) {
          return pageName;
        }
      }
    }
  }

  // Fallback: Try to get from document title (remove " - Power BI" suffix)
  if (document.title) {
    const cleanTitle = document.title.replace(/ - Power BI.*$/i, '').trim();
    if (cleanTitle && cleanTitle.length > 5) {
      return cleanTitle;
    }
  }

  // Otherwise use URL path
  const path = window.location.pathname;
  const parts = path.split("/").filter((p) => p);

  if (parts.length > 0) {
    return parts[parts.length - 1];
  }

  return 'Power BI Report';
}

// Create annotation element from annotation data
// [Fix #1] Uses appendChild instead of innerHTML += to preserve SVG namespace
// [Fix #2] Uses startPoint/endPoint for correct arrow/line direction
function createAnnotationElement(annotation, number) {
  const box = document.createElement("div");
  box.className = "pbi-annotation-box";
  box.dataset.id = annotation.id;
  box.style.left = annotation.x + "px";
  box.style.top = annotation.y + "px";
  box.style.width = annotation.width + "px";
  box.style.height = annotation.height + "px";

  const tool = annotation.tool || 'rectangle';
  const color = annotation.color || '#0078d4';
  box.style.borderColor = color;

  // Helper to create and append the number badge via DOM (not innerHTML)
  function appendBadge() {
    const badge = document.createElement('div');
    badge.className = 'pbi-annotation-number';
    badge.textContent = number;
    box.appendChild(badge);
  }

  // [Fix #2] Determine direction from stored start/end points (backward-compatible)
  function getDirection() {
    if (annotation.startPoint && annotation.endPoint) {
      return {
        x1: annotation.startPoint.x <= annotation.endPoint.x ? 0 : annotation.width,
        y1: annotation.startPoint.y <= annotation.endPoint.y ? 0 : annotation.height,
        x2: annotation.startPoint.x <= annotation.endPoint.x ? annotation.width : 0,
        y2: annotation.startPoint.y <= annotation.endPoint.y ? annotation.height : 0,
      };
    }
    // Backward compat: default to top-left → bottom-right
    return { x1: 0, y1: 0, x2: annotation.width, y2: annotation.height };
  }

  // Render based on tool type
  if (tool === 'freehand' && annotation.freehandPath) {
    const points = annotation.freehandPath;
    const minX = Math.min(...points.map(p => p.x));
    const minY = Math.min(...points.map(p => p.y));

    const pathData = points.map((p, i) => {
      const x = p.x - minX;
      const y = p.y - minY;
      return i === 0 ? `M ${x} ${y}` : `L ${x} ${y}`;
    }).join(' ');

    box.innerHTML = `<svg class="pbi-freehand-svg" width="${annotation.width}" height="${annotation.height}" style="position: absolute; top: 0; left: 0; pointer-events: none;"><path d="${pathData}" stroke="${color}" stroke-width="3" fill="none"/></svg>`;
    appendBadge();
  } else if (tool === 'arrow') {
    const dir = getDirection();
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute('width', annotation.width);
    svg.setAttribute('height', annotation.height);
    svg.style.position = 'absolute';
    svg.style.top = '0';
    svg.style.left = '0';
    svg.style.pointerEvents = 'none';

    const headLen = 20;
    const angle = Math.atan2(dir.y2 - dir.y1, dir.x2 - dir.x1);
    const arrowX1 = dir.x2 - headLen * Math.cos(angle - Math.PI / 6);
    const arrowY1 = dir.y2 - headLen * Math.sin(angle - Math.PI / 6);
    const arrowX2 = dir.x2 - headLen * Math.cos(angle + Math.PI / 6);
    const arrowY2 = dir.y2 - headLen * Math.sin(angle + Math.PI / 6);

    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute('d', `M ${dir.x1} ${dir.y1} L ${dir.x2} ${dir.y2} M ${arrowX1} ${arrowY1} L ${dir.x2} ${dir.y2} L ${arrowX2} ${arrowY2}`);
    path.setAttribute('stroke', color);
    path.setAttribute('stroke-width', '3');
    path.setAttribute('fill', 'none');

    svg.appendChild(path);
    box.appendChild(svg);
    appendBadge();
  } else if (tool === 'line') {
    const dir = getDirection();
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute('width', annotation.width);
    svg.setAttribute('height', annotation.height);
    svg.style.position = 'absolute';
    svg.style.top = '0';
    svg.style.left = '0';
    svg.style.pointerEvents = 'none';

    const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
    line.setAttribute('x1', dir.x1);
    line.setAttribute('y1', dir.y1);
    line.setAttribute('x2', dir.x2);
    line.setAttribute('y2', dir.y2);
    line.setAttribute('stroke', color);
    line.setAttribute('stroke-width', '3');

    svg.appendChild(line);
    box.appendChild(svg);
    appendBadge();
  } else if (tool === 'circle') {
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute('width', annotation.width);
    svg.setAttribute('height', annotation.height);
    svg.style.position = 'absolute';
    svg.style.top = '0';
    svg.style.left = '0';
    svg.style.pointerEvents = 'none';

    const ellipse = document.createElementNS("http://www.w3.org/2000/svg", "ellipse");
    ellipse.setAttribute('cx', annotation.width / 2);
    ellipse.setAttribute('cy', annotation.height / 2);
    ellipse.setAttribute('rx', annotation.width / 2);
    ellipse.setAttribute('ry', annotation.height / 2);
    ellipse.setAttribute('stroke', color);
    ellipse.setAttribute('stroke-width', '3');
    ellipse.setAttribute('fill', 'none');

    svg.appendChild(ellipse);
    box.appendChild(svg);
    appendBadge();
  } else {
    // Rectangle - default
    appendBadge();
  }

  box.addEventListener("click", (e) => {
    e.stopPropagation();
    showAnnotationComment(annotation.id);
  });

  return box;
}

// Save annotations to Chrome storage [Fix #4] - uses cache to avoid read-modify-write race
function saveAnnotations() {
  const pageKey = getPageKey();

  if (allAnnotationsCache === null) {
    // Cache not loaded yet - fall back to read-modify-write
    chrome.storage.local.get(["annotations"], (result) => {
      if (chrome.runtime.lastError) {
        console.error("Failed to load annotations for saving:", chrome.runtime.lastError);
        return;
      }
      allAnnotationsCache = result.annotations || {};
      allAnnotationsCache[pageKey] = annotations;
      chrome.storage.local.set({ annotations: allAnnotationsCache }, () => {
        if (chrome.runtime.lastError) {
          console.error("Failed to save annotations:", chrome.runtime.lastError);
        }
      });
    });
  } else {
    // Cache is available - write directly (no read needed)
    allAnnotationsCache[pageKey] = annotations;
    chrome.storage.local.set({ annotations: allAnnotationsCache }, () => {
      if (chrome.runtime.lastError) {
        console.error("Failed to save annotations:", chrome.runtime.lastError);
      }
    });
  }
}

// Load annotations from Chrome storage [Fix #4, #5]
function loadAnnotations() {
  const pageKey = getPageKey();
  const legacyKey = window.location.pathname; // [Fix #5] Old key format for migration

  chrome.storage.local.get(["annotations"], (result) => {
    if (chrome.runtime.lastError) {
      console.error("Failed to load annotations:", chrome.runtime.lastError);
      return;
    }
    allAnnotationsCache = result.annotations || {};
    annotations = allAnnotationsCache[pageKey] || [];

    // [Fix #5] Migrate from legacy pathname-only key if new key has no data
    if (annotations.length === 0 && legacyKey !== pageKey && allAnnotationsCache[legacyKey]) {
      annotations = allAnnotationsCache[legacyKey];
      allAnnotationsCache[pageKey] = annotations;
      delete allAnnotationsCache[legacyKey];
      chrome.storage.local.set({ annotations: allAnnotationsCache });
    }

    // Render saved annotations on page
    annotations.forEach((annotation, index) => {
      const box = createAnnotationElement(annotation, index + 1);
      document.body.appendChild(box);
    });

    renderComments();
    renderPageList();
  });
}

// Get unique key for current page [Fix #5] - includes query params for Power BI page navigation
function getPageKey() {
  return window.location.pathname + window.location.search;
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
