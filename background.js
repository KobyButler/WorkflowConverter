chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === 'setServerEndpoint') {
      chrome.storage.sync.set({ serverEndpoint: message.endpoint }, () => {
        sendResponse({ status: 'success' });
      });
    } else if (message.action === 'getServerEndpoint') {
      chrome.storage.sync.get({ serverEndpoint: 'http://your-default-endpoint/upload' }, result => {
        sendResponse({ endpoint: result.serverEndpoint });
      });
    }
    return true; // Needed to keep the message channel open
  });
  