const output = document.getElementById("output");
const statusEl = document.getElementById("status");
const refreshButton = document.getElementById("refresh");
const copyButton = document.getElementById("copy");

const setStatus = (text) => {
  statusEl.textContent = text;
};

const setOutput = (text) => {
  output.value = text;
  copyButton.disabled = text.trim().length === 0;
};

const fetchDebugInfo = async () => {
  setStatus("取得中...");
  setOutput("");

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab || typeof tab.id !== "number") {
    setStatus("アクティブなタブが見つかりません。");
    return;
  }

  try {
    const response = await chrome.tabs.sendMessage(tab.id, { type: "OCE_DEBUG" });
    if (!response) {
      setStatus("デバッグ情報が取得できませんでした。");
      return;
    }
    setOutput(JSON.stringify(response, null, 2));
    setStatus("取得完了");
  } catch (error) {
    setStatus(`取得失敗: ${error?.message || error}`);
  }
};

refreshButton.addEventListener("click", () => {
  void fetchDebugInfo();
});

copyButton.addEventListener("click", async () => {
  try {
    await navigator.clipboard.writeText(output.value || "");
    setStatus("コピーしました");
  } catch (error) {
    setStatus("コピーに失敗しました");
  }
});

void fetchDebugInfo();
