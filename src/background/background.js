// Power BI Annotator - Background Service Worker

// This service worker handles extension installation, updates,
// and screenshot capture via the activeTab permission flow.

// Tracks when a screenshot capture is pending (waiting for icon click).
// Also persisted to chrome.storage.session so it survives service worker restarts.
let pendingCapture = null; // { tabId: number } or null

chrome.runtime.onInstalled.addListener((details) => {
  if (details.reason === 'install') {
    console.log('Power BI Annotator installed');

    // Open a welcome page or show notification
    chrome.tabs.create({
      url: 'https://app.powerbi.com/'
    });
  } else if (details.reason === 'update') {
    console.log('Power BI Annotator updated');
  }
});

// Handle extension icon clicks.
// If a screenshot capture is pending, capture it (activeTab is granted by the click).
// Otherwise, toggle the sidebar as usual.
chrome.action.onClicked.addListener((tab) => {
  // Fast path: in-memory state is available (service worker stayed alive)
  if (pendingCapture && pendingCapture.tabId === tab.id) {
    captureAndSend(tab);
    return;
  }

  // Slow path: service worker may have restarted, check session storage
  chrome.storage.session.get('pendingCapture', (result) => {
    if (result.pendingCapture && result.pendingCapture.tabId === tab.id) {
      pendingCapture = result.pendingCapture;
      captureAndSend(tab);
    } else {
      // Normal behavior: toggle sidebar
      chrome.tabs.sendMessage(tab.id, { action: 'toggleSidebar' });
    }
  });
});

// Hide the instruction modal, wait for DOM update, then capture screenshot
function captureAndSend(tab) {
  chrome.tabs.sendMessage(tab.id, { action: 'hideForCapture' }, () => {
    setTimeout(() => {
      chrome.tabs.captureVisibleTab(tab.windowId, { format: 'png' }, (dataUrl) => {
        if (chrome.runtime.lastError) {
          console.error('Screenshot capture failed:', chrome.runtime.lastError.message);
          chrome.tabs.sendMessage(tab.id, {
            action: 'screenshotResult',
            screenshot: null,
            error: chrome.runtime.lastError.message
          });
        } else {
          chrome.tabs.sendMessage(tab.id, {
            action: 'screenshotResult',
            screenshot: dataUrl
          });
        }
        clearPendingCapture();
      });
    }, 200);
  });
}

function setPendingCapture(tabId) {
  pendingCapture = { tabId: tabId };
  chrome.storage.session.set({ pendingCapture: pendingCapture });
  chrome.action.setBadgeText({ text: '📸' });
  chrome.action.setBadgeBackgroundColor({ color: '#0078d4' });
}

function clearPendingCapture() {
  pendingCapture = null;
  chrome.storage.session.remove('pendingCapture');
  chrome.action.setBadgeText({ text: '' });
}

// Handle messages from content script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'prepareCapture') {
    setPendingCapture(sender.tab.id);
    sendResponse({ ok: true });
  } else if (request.action === 'cancelCapture') {
    clearPendingCapture();
    sendResponse({ ok: true });
  } else if (request.action === 'captureForCache') {
    // Silent screenshot capture using host_permissions (no user gesture needed)
    chrome.tabs.captureVisibleTab(sender.tab.windowId, { format: 'png' }, (dataUrl) => {
      if (chrome.runtime.lastError) {
        sendResponse({ screenshot: null, error: chrome.runtime.lastError.message });
      } else {
        sendResponse({ screenshot: dataUrl });
      }
    });
    return true; // Keep message channel open for async captureVisibleTab
  }
});
