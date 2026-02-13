// Power BI Page Context Script
// This script runs in the page context (not content script isolated world)
// to access the Power BI embed API (powerbi.embeds) which is only available
// to page-level JavaScript, not content scripts.
//
// Also intercepts history.pushState/replaceState for instant SPA navigation
// detection. Content scripts can't override these in their isolated world,
// so we do it here in the page world and notify via window.postMessage.

(function() {
  var PBI_MSG_TYPE = '__pbi_annotator_page_info__';
  var PBI_REQUEST_TYPE = '__pbi_annotator_request_page_info__';
  var NAV_MSG_TYPE = '__pbi_annotator_navigation__';

  // --- Instant SPA navigation detection ---
  // Override pushState/replaceState in the page world so the content script
  // can detect URL changes immediately instead of relying on polling.
  // Uses window.postMessage (proven to cross the isolated world boundary).
  var originalPushState = history.pushState;
  var originalReplaceState = history.replaceState;

  history.pushState = function() {
    originalPushState.apply(this, arguments);
    window.postMessage({ type: NAV_MSG_TYPE }, '*');
  };

  history.replaceState = function() {
    originalReplaceState.apply(this, arguments);
    window.postMessage({ type: NAV_MSG_TYPE }, '*');
  };
  var lastSentUrl = '';
  var lastSentName = '';
  var retryCount = 0;
  var maxRetries = 30; // Try for 30 seconds at 1s intervals

  function getActivePageInfo(forceUpdate) {
    try {
      // Check if Power BI embed API is available
      if (typeof powerbi === 'undefined' || !powerbi.embeds || powerbi.embeds.length === 0) {
        retryCount++;
        return false; // API not ready
      }

      var report = powerbi.embeds[0];
      if (!report || typeof report.getActivePage !== 'function') {
        return false;
      }

      var currentUrl = window.location.pathname + window.location.search;

      report.getActivePage().then(function(page) {
        if (page && page.displayName) {
          // Send update if page name or URL changed, or if forced
          if (forceUpdate || page.displayName !== lastSentName || currentUrl !== lastSentUrl) {
            lastSentName = page.displayName;
            lastSentUrl = currentUrl;

            window.postMessage({
              type: PBI_MSG_TYPE,
              displayName: page.displayName,
              name: page.name,
              url: currentUrl
            }, '*');
          }
        }
      }).catch(function() {
        // Silently ignore — API may not be ready yet
      });

      return true; // API was found
    } catch (e) {
      return false;
    }
  }

  // Listen for on-demand requests from the content script (triggered on page navigation)
  window.addEventListener('message', function(event) {
    if (event.source !== window) return;
    if (!event.data || event.data.type !== PBI_REQUEST_TYPE) return;
    getActivePageInfo(true);
  });

  // Poll until the API is found, then stop.
  // The content script sends on-demand requests on page navigation, so continuous
  // polling is not needed once the API is available.
  var pollInterval = setInterval(function() {
    var found = getActivePageInfo(false);
    if (found || retryCount >= maxRetries) {
      clearInterval(pollInterval);
    }
  }, 1000);

  // Also try immediately
  getActivePageInfo(false);
})();
