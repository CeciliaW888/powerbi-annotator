// Power BI Page Context Script
// This script runs in the page context (not content script isolated world)
// to access the Power BI embed API (powerbi.embeds) which is only available
// to page-level JavaScript, not content scripts.

(function() {
  var PBI_MSG_TYPE = '__pbi_annotator_page_info__';
  var PBI_REQUEST_TYPE = '__pbi_annotator_request_page_info__';
  var lastSentUrl = '';
  var lastSentName = '';
  var retryCount = 0;
  var maxRetries = 30; // Try for 30 seconds
  var hasLoggedStatus = false;

  function getActivePageInfo(forceUpdate) {
    try {
      // Check if Power BI embed API is available
      if (typeof powerbi === 'undefined') {
        if (!hasLoggedStatus) {
          console.log('[PBI Embed API] powerbi object not found - API not available in this context');
          hasLoggedStatus = true;
        }
        if (retryCount < maxRetries) {
          retryCount++;
        }
        return;
      }
      
      if (!powerbi.embeds || powerbi.embeds.length === 0) {
        if (!hasLoggedStatus) {
          console.log('[PBI Embed API] powerbi.embeds is empty - no embedded reports detected');
          hasLoggedStatus = true;
        }
        if (retryCount < maxRetries) {
          retryCount++;
        }
        return;
      }
      
      var report = powerbi.embeds[0];
      console.log('[PBI Embed API] Found embed, type:', report.embedtype || 'unknown');
      
      if (!report || typeof report.getActivePage !== 'function') {
        console.log('[PBI Embed API] getActivePage method not available on report object');
        return;
      }
      
      var currentUrl = window.location.pathname + window.location.search;
      
      report.getActivePage().then(function(page) {
        if (page && page.displayName) {
          console.log('[PBI Embed API] Successfully got active page:', {
            displayName: page.displayName,
            name: page.name,
            url: currentUrl
          });
          
          // Send update if page name or URL changed, or if forced
          if (forceUpdate || page.displayName !== lastSentName || currentUrl !== lastSentUrl) {
            lastSentName = page.displayName;
            lastSentUrl = currentUrl;
            
            console.log('[PBI Embed API] Sending page info to content script');
            
            window.postMessage({
              type: PBI_MSG_TYPE,
              displayName: page.displayName,
              name: page.name,
              url: currentUrl
            }, '*');
          }
          
          // Successfully got page info, reset retry count
          retryCount = 0;
          hasLoggedStatus = false;
        } else {
          console.log('[PBI Embed API] Page object missing displayName:', page);
        }
      }).catch(function(err) {
        console.log('[PBI Embed API] Error getting active page:', err);
      });
    } catch (e) {
      console.log('[PBI Embed API] Exception in getActivePageInfo:', e);
    }
  }

  // Listen for requests to update page info immediately
  window.addEventListener('message', function(event) {
    if (event.source !== window) return;
    if (!event.data || event.data.type !== PBI_REQUEST_TYPE) return;
    
    console.log('[PBI Embed API] Received request for page info update');
    getActivePageInfo(true); // Force update
  });

  // Poll periodically — Power BI may take time to initialize embeds
  // Use more frequent polling initially, then slow down
  var pollInterval = setInterval(function() {
    getActivePageInfo(false);
    
    // After successful detection or max retries, slow down polling to reduce CPU
    if (retryCount === 0 || retryCount >= maxRetries) {
      clearInterval(pollInterval);
      console.log('[PBI Embed API] Switching to slower polling (2s intervals)');
      setInterval(getActivePageInfo, 2000); // Poll every 2 seconds after initialization
    }
  }, 500); // Poll every 500ms initially for fast startup
  
  // Also try immediately
  getActivePageInfo(false);
  
  console.log('[PBI Embed API] Page info script injected and running');
})();
