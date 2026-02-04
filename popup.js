const output = document.getElementById("output");
const statusEl = document.getElementById("status");
const refreshButton = document.getElementById("refresh");
const copyButton = document.getElementById("copy");
const contactsInput = document.getElementById("contacts-input");
const contactsSaveButton = document.getElementById("contacts-save");
const contactsClearButton = document.getElementById("contacts-clear");
const contactsStatus = document.getElementById("contacts-status");
const contactsCount = document.getElementById("contacts-count");
const CONTACTS_STORAGE_KEY = "oceContacts";

const setStatus = (text) => {
  statusEl.textContent = text;
};

const setOutput = (text) => {
  output.value = text;
  copyButton.disabled = text.trim().length === 0;
};

const setContactsStatus = (text) => {
  contactsStatus.textContent = text;
};

const setContactsCount = (count) => {
  contactsCount.textContent = String(count);
};

const detectDelimiter = (lines) => {
  if (lines.some((line) => line.includes("\t"))) return "\t";
  if (lines.some((line) => line.includes(","))) return ",";
  if (lines.some((line) => line.includes(";"))) return ";";
  return null;
};

const splitLine = (line, delimiter) => {
  const result = [];
  let current = "";
  let inQuotes = false;
  let i = 0;
  while (i < line.length) {
    const char = line[i];
    if (char === "\"") {
      const next = line[i + 1];
      if (inQuotes && next === "\"") {
        current += "\"";
        i += 2;
        continue;
      }
      inQuotes = !inQuotes;
      i += 1;
      continue;
    }
    if (!inQuotes && char === delimiter) {
      result.push(current);
      current = "";
      i += 1;
      continue;
    }
    current += char;
    i += 1;
  }
  result.push(current);
  return result.map((value) =>
    value.trim().replaceAll(/(^"|"$)/g, "")
  );
};

const looksLikeEmail = (value) => /@/.test(value);

const detectHeaderIndices = (fields) => {
  const lowered = fields.map((value) => value.toLowerCase());
  const emailIndex = lowered.findIndex(
    (value) => value.includes("email") || value.includes("mail")
  );
  const nameIndex = lowered.findIndex(
    (value) =>
      value.includes("name") ||
      value.includes("full") ||
      value.includes("氏名") ||
      value.includes("名前")
  );
  if (emailIndex >= 0 && nameIndex >= 0) {
    return { emailIndex, nameIndex, isHeader: true };
  }
  return null;
};

const getNonEmptyLines = (text) =>
  text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

const guessColumnIndices = (fields) => {
  if (looksLikeEmail(fields[0]) && fields[1]) {
    return { emailIndex: 0, nameIndex: 1 };
  }
  if (looksLikeEmail(fields[1])) {
    return { emailIndex: 1, nameIndex: 0 };
  }
  return { emailIndex: 1, nameIndex: 0 };
};

const resolveColumnIndices = (lines, delimiter) => {
  const header = detectHeaderIndices(splitLine(lines[0], delimiter));
  if (header?.isHeader) {
    return {
      nameIndex: header.nameIndex,
      emailIndex: header.emailIndex,
      startIndex: 1
    };
  }
  const firstRow = splitLine(lines[0], delimiter);
  const guess = guessColumnIndices(firstRow);
  return { ...guess, startIndex: 0 };
};

const parseContactsFromLines = (lines, delimiter, nameIndex, emailIndex, startIndex) => {
  const seen = new Set();
  const contacts = [];
  let skipped = 0;

  for (let i = startIndex; i < lines.length; i += 1) {
    const fields = splitLine(lines[i], delimiter);
    const name = (fields[nameIndex] || "").trim();
    const email = (fields[emailIndex] || "").trim();
    if (!name || !email || !looksLikeEmail(email)) {
      skipped += 1;
      continue;
    }
    const key = `${name}\n${email}`.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    contacts.push({ name, email });
  }

  return { contacts, skipped };
};

const parseContacts = (text) => {
  const lines = getNonEmptyLines(text);
  if (lines.length === 0) {
    return { contacts: [], skipped: 0, error: "入力が空です。" };
  }
  const delimiter = detectDelimiter(lines);
  if (!delimiter) {
    return {
      contacts: [],
      skipped: lines.length,
      error: "区切り文字(タブ/カンマ/セミコロン)が見つかりません。"
    };
  }

  const { nameIndex, emailIndex, startIndex } = resolveColumnIndices(lines, delimiter);
  const { contacts, skipped } = parseContactsFromLines(
    lines,
    delimiter,
    nameIndex,
    emailIndex,
    startIndex
  );
  return { contacts, skipped, error: "" };
};

const contactsToText = (contacts) => {
  const header = "full_name\temail_address";
  const lines = contacts.map((entry) => `${entry.name}\t${entry.email}`);
  return [header, ...lines].join("\n");
};

const loadContactsCount = () => {
  if (!chrome?.storage?.local) return;
  chrome.storage.local.get(CONTACTS_STORAGE_KEY, (result) => {
    const list = result[CONTACTS_STORAGE_KEY];
    const contacts = Array.isArray(list) ? list : [];
    setContactsCount(contacts.length);
    if (contacts.length > 0) {
      contactsInput.value = contactsToText(contacts);
    }
  });
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
    setStatus(`コピーに失敗しました: ${error?.message || error}`);
  }
});

contactsSaveButton.addEventListener("click", () => {
  setContactsStatus("");
  const text = contactsInput.value || "";
  const { contacts, skipped, error } = parseContacts(text);
  if (error) {
    setContactsStatus(error);
    return;
  }
  if (!chrome?.storage?.local) {
    setContactsStatus("保存先が利用できません。");
    return;
  }
  chrome.storage.local.set({ [CONTACTS_STORAGE_KEY]: contacts }, () => {
    setContactsCount(contacts.length);
    contactsInput.value = contactsToText(contacts);
    setContactsStatus(
      `保存しました (${contacts.length}件, スキップ${skipped}件)`
    );
  });
});

contactsClearButton.addEventListener("click", () => {
  contactsInput.value = "";
  if (!chrome?.storage?.local) {
    setContactsStatus("保存先が利用できません。");
    return;
  }
  chrome.storage.local.remove(CONTACTS_STORAGE_KEY, () => {
    setContactsCount(0);
    setContactsStatus("クリアしました");
  });
});

loadContactsCount();
void fetchDebugInfo();
