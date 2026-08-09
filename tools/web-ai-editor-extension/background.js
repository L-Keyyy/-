chrome.action.onClicked.addListener(async (tab) => {
  if (!tab.id || !tab.url) return;

  const supported = /^(https?:|file:)/.test(tab.url);
  if (!supported) {
    await chrome.action.setBadgeText({ tabId: tab.id, text: "!" });
    await chrome.action.setBadgeBackgroundColor({
      tabId: tab.id,
      color: "#dc4c56",
    });
    return;
  }

  try {
    await chrome.scripting.insertCSS({
      target: { tabId: tab.id },
      files: ["content.css"],
    });
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ["content.js"],
    });
    await chrome.action.setBadgeText({ tabId: tab.id, text: "EDIT" });
    await chrome.action.setBadgeBackgroundColor({
      tabId: tab.id,
      color: "#2563eb",
    });
  } catch (error) {
    console.error("Web AI 页面修改器启动失败", error);
    await chrome.action.setBadgeText({ tabId: tab.id, text: "!" });
    await chrome.action.setBadgeBackgroundColor({
      tabId: tab.id,
      color: "#dc4c56",
    });
  }
});

chrome.runtime.onMessage.addListener((message, sender) => {
  if (message?.type !== "WEB_AI_EDITOR_CLOSED" || !sender.tab?.id) return;
  chrome.action.setBadgeText({ tabId: sender.tab.id, text: "" });
});
