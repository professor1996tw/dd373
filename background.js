// SunWater DD373 自動進貨系統 - Background Service Worker
// 點擊擴充功能圖示時，開啟主介面頁面（全視窗模式）

chrome.action.onClicked.addListener((tab) => {
  chrome.tabs.create({
    url: chrome.runtime.getURL('index.html')
  });
});
