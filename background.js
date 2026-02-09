// Luma Voice Assistant — Background Service Worker

chrome.runtime.onInstalled.addListener((details) => {
  if (details.reason === "install") {
    chrome.tabs.create({ url: "demo.html" });
  }
});

// Toggle assistant via keyboard shortcut (Ctrl+Shift+S)
chrome.commands.onCommand.addListener((command) => {
  if (command === "toggle-assistant") {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      if (tabs[0]) {
        chrome.tabs.sendMessage(tabs[0].id, { action: "toggle" });
      }
    });
  }
});

// Open demo page when extension icon is clicked
chrome.action.onClicked.addListener(() => {
  chrome.tabs.create({ url: "demo.html" });
});
