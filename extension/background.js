// Clicking the toolbar icon opens the side panel (it stays open while you work on the claim site).
chrome.runtime.onInstalled.addListener(() => {
  chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(() => {});
});
