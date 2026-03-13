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
const EXCLUDE_KEYWORDS_KEY = "oceMeetingExcludeKeywords";
const WORK_HOURS_KEY = "oceMeetingWorkHours";
const ROOM_EMAIL_DOMAIN_KEY = "oceRoomEmailDomain";
const COLOR_THRESHOLDS_KEY = "oceColorThresholds";
const EXCLUDE_RULES_KEY = "oceMeetingExcludeRules";
const excludeExactInput = document.getElementById("exclude-exact");
const excludePrefixInput = document.getElementById("exclude-prefix");
const excludeSuffixInput = document.getElementById("exclude-suffix");
const excludePartialInput = document.getElementById("exclude-partial");
const excludeRulesSaveButton = document.getElementById("exclude-rules-save");
const excludeRulesClearButton = document.getElementById("exclude-rules-clear");
const excludeRulesStatus = document.getElementById("exclude-rules-status");
const thresholdLowInput = document.getElementById("threshold-low");
const thresholdMidInput = document.getElementById("threshold-mid");
const thresholdHighInput = document.getElementById("threshold-high");
const thresholdSaveButton = document.getElementById("threshold-save");
const thresholdResetButton = document.getElementById("threshold-reset");
const thresholdStatus = document.getElementById("threshold-status");
const excludeInput = document.getElementById("exclude-keywords-input");
const excludeSaveButton = document.getElementById("exclude-keywords-save");
const excludeClearButton = document.getElementById("exclude-keywords-clear");
const excludeStatus = document.getElementById("exclude-keywords-status");
const roomDomainInput = document.getElementById("room-email-domain-input");
const roomDomainSaveButton = document.getElementById("room-email-domain-save");
const roomDomainClearButton = document.getElementById("room-email-domain-clear");
const roomDomainStatus = document.getElementById("room-email-domain-status");
const workStartInput = document.getElementById("work-start");
const workEndInput = document.getElementById("work-end");
const workLunchInput = document.getElementById("work-lunch");
const workHoursSaveButton = document.getElementById("work-hours-save");
const workHoursResetButton = document.getElementById("work-hours-reset");
const workHoursStatus = document.getElementById("work-hours-status");

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
  const idIndex = lowered.findIndex(
    (value) =>
      value === "id" ||
      value.includes("slack") ||
      value.includes("handle")
  );
  if (emailIndex >= 0 && nameIndex >= 0) {
    return { emailIndex, nameIndex, idIndex, isHeader: true };
  }
  return null;
};

const getNonEmptyLines = (text) =>
  text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

const guessColumnIndices = (fields) => {
  const emailIndex = fields.findIndex((value) => looksLikeEmail(value));
  const resolvedEmailIndex = emailIndex >= 0 ? emailIndex : 1;
  const nameIndex =
    resolvedEmailIndex === 0 ? 1 : 0;
  let idIndex = -1;
  if (fields.length >= 3) {
    for (let i = 0; i < fields.length; i += 1) {
      if (i !== resolvedEmailIndex && i !== nameIndex) {
        idIndex = i;
        break;
      }
    }
  }
  return { emailIndex: resolvedEmailIndex, nameIndex, idIndex };
};

const resolveColumnIndices = (lines, delimiter) => {
  const header = detectHeaderIndices(splitLine(lines[0], delimiter));
  if (header?.isHeader) {
    return {
      nameIndex: header.nameIndex,
      emailIndex: header.emailIndex,
      idIndex: header.idIndex ?? -1,
      startIndex: 1
    };
  }
  const firstRow = splitLine(lines[0], delimiter);
  const guess = guessColumnIndices(firstRow);
  return { ...guess, startIndex: 0 };
};

const parseContactsFromLines = (
  lines,
  delimiter,
  nameIndex,
  emailIndex,
  idIndex,
  startIndex
) => {
  const seen = new Set();
  const contacts = [];
  let skipped = 0;

  for (let i = startIndex; i < lines.length; i += 1) {
    const fields = splitLine(lines[i], delimiter);
    const name = (fields[nameIndex] || "").trim();
    const email = (fields[emailIndex] || "").trim();
    const id = idIndex >= 0 ? (fields[idIndex] || "").trim() : "";
    if (!name || !email || !looksLikeEmail(email)) {
      skipped += 1;
      continue;
    }
    const key = `${name}\n${email}`.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    contacts.push({ name, email, id });
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

  const { nameIndex, emailIndex, idIndex, startIndex } = resolveColumnIndices(
    lines,
    delimiter
  );
  const { contacts, skipped } = parseContactsFromLines(
    lines,
    delimiter,
    nameIndex,
    emailIndex,
    idIndex,
    startIndex
  );
  return { contacts, skipped, error: "" };
};

const contactsToText = (contacts) => {
  const header = "full_name\temail_address\tid";
  const lines = contacts.map((entry) =>
    `${entry.name}\t${entry.email}\t${entry.id || ""}`
  );
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

const loadExcludeKeywords = () => {
  if (!chrome?.storage?.local) return;
  chrome.storage.local.get(EXCLUDE_KEYWORDS_KEY, (result) => {
    const keywords = result[EXCLUDE_KEYWORDS_KEY];
    if (Array.isArray(keywords) && keywords.length > 0) {
      excludeInput.value = keywords.join(", ");
    }
  });
};

excludeSaveButton.addEventListener("click", () => {
  const text = excludeInput.value || "";
  const keywords = text
    .split(",")
    .map((kw) => kw.trim())
    .filter((kw) => kw.length > 0);
  if (!chrome?.storage?.local) {
    excludeStatus.textContent = "保存先が利用できません。";
    return;
  }
  chrome.storage.local.set({ [EXCLUDE_KEYWORDS_KEY]: keywords }, () => {
    excludeStatus.textContent = `保存しました (${keywords.length}件)`;
  });
});

excludeClearButton.addEventListener("click", () => {
  excludeInput.value = "";
  if (!chrome?.storage?.local) {
    excludeStatus.textContent = "保存先が利用できません。";
    return;
  }
  chrome.storage.local.remove(EXCLUDE_KEYWORDS_KEY, () => {
    excludeStatus.textContent = "クリアしました";
  });
});

const parseCommaSeparated = (text) =>
  (text || "").split(",").map((s) => s.trim()).filter((s) => s.length > 0);

const loadExcludeRules = () => {
  if (!chrome?.storage?.local) return;
  chrome.storage.local.get(EXCLUDE_RULES_KEY, (result) => {
    const r = result[EXCLUDE_RULES_KEY];
    if (r) {
      excludeExactInput.value = (r.exact || []).join(", ");
      excludePrefixInput.value = (r.prefix || []).join(", ");
      excludeSuffixInput.value = (r.suffix || []).join(", ");
      excludePartialInput.value = (r.partial || []).join(", ");
    }
  });
};

excludeRulesSaveButton.addEventListener("click", () => {
  const rules = {
    exact: parseCommaSeparated(excludeExactInput.value),
    prefix: parseCommaSeparated(excludePrefixInput.value),
    suffix: parseCommaSeparated(excludeSuffixInput.value),
    partial: parseCommaSeparated(excludePartialInput.value)
  };
  const total = rules.exact.length + rules.prefix.length + rules.suffix.length + rules.partial.length;
  if (!chrome?.storage?.local) {
    excludeRulesStatus.textContent = "保存先が利用できません。";
    return;
  }
  chrome.storage.local.set({ [EXCLUDE_RULES_KEY]: rules }, () => {
    excludeRulesStatus.textContent = `保存しました (${total}件)`;
  });
});

excludeRulesClearButton.addEventListener("click", () => {
  excludeExactInput.value = "";
  excludePrefixInput.value = "";
  excludeSuffixInput.value = "";
  excludePartialInput.value = "";
  if (!chrome?.storage?.local) {
    excludeRulesStatus.textContent = "保存先が利用できません。";
    return;
  }
  chrome.storage.local.remove(EXCLUDE_RULES_KEY, () => {
    excludeRulesStatus.textContent = "クリアしました";
  });
});

const loadRoomEmailDomain = () => {
  if (!chrome?.storage?.local) return;
  chrome.storage.local.get(ROOM_EMAIL_DOMAIN_KEY, (result) => {
    const domain = result[ROOM_EMAIL_DOMAIN_KEY];
    if (domain) {
      roomDomainInput.value = domain;
    }
  });
};

roomDomainSaveButton.addEventListener("click", () => {
  const domain = roomDomainInput.value.trim();
  if (!chrome?.storage?.local) {
    roomDomainStatus.textContent = "保存先が利用できません。";
    return;
  }
  chrome.storage.local.set({ [ROOM_EMAIL_DOMAIN_KEY]: domain }, () => {
    roomDomainStatus.textContent = domain
      ? `保存しました: ${domain}`
      : "空で保存しました";
  });
});

roomDomainClearButton.addEventListener("click", () => {
  roomDomainInput.value = "";
  if (!chrome?.storage?.local) {
    roomDomainStatus.textContent = "保存先が利用できません。";
    return;
  }
  chrome.storage.local.remove(ROOM_EMAIL_DOMAIN_KEY, () => {
    roomDomainStatus.textContent = "クリアしました";
  });
});

const loadWorkHours = () => {
  if (!chrome?.storage?.local) return;
  chrome.storage.local.get(WORK_HOURS_KEY, (result) => {
    const wh = result[WORK_HOURS_KEY];
    if (wh) {
      workStartInput.value = wh.start ?? 9;
      workEndInput.value = wh.end ?? 18;
      workLunchInput.value = wh.lunch ?? 60;
    }
  });
};

workHoursSaveButton.addEventListener("click", () => {
  const start = Number.parseInt(workStartInput.value, 10);
  const end = Number.parseInt(workEndInput.value, 10);
  const lunch = Number.parseInt(workLunchInput.value, 10);
  if (Number.isNaN(start) || Number.isNaN(end) || start >= end) {
    workHoursStatus.textContent = "開始 < 終了 で入力してください。";
    return;
  }
  const lunchVal = Number.isNaN(lunch) ? 60 : Math.max(0, lunch);
  if (!chrome?.storage?.local) {
    workHoursStatus.textContent = "保存先が利用できません。";
    return;
  }
  chrome.storage.local.set({ [WORK_HOURS_KEY]: { start, end, lunch: lunchVal } }, () => {
    const effective = end - start - lunchVal / 60;
    workHoursStatus.textContent = `保存しました (${start}:00〜${end}:00, 実働${effective}h/日)`;
  });
});

workHoursResetButton.addEventListener("click", () => {
  workStartInput.value = 9;
  workEndInput.value = 18;
  workLunchInput.value = 60;
  if (!chrome?.storage?.local) {
    workHoursStatus.textContent = "保存先が利用できません。";
    return;
  }
  chrome.storage.local.remove(WORK_HOURS_KEY, () => {
    workHoursStatus.textContent = "リセットしました (9:00〜18:00)";
  });
});

const loadColorThresholds = () => {
  if (!chrome?.storage?.local) return;
  chrome.storage.local.get(COLOR_THRESHOLDS_KEY, (result) => {
    const t = result[COLOR_THRESHOLDS_KEY];
    if (t) {
      thresholdLowInput.value = t.low ?? 50;
      thresholdMidInput.value = t.mid ?? 70;
      thresholdHighInput.value = t.high ?? 90;
    }
  });
};

thresholdSaveButton.addEventListener("click", () => {
  const low = Number.parseInt(thresholdLowInput.value, 10);
  const mid = Number.parseInt(thresholdMidInput.value, 10);
  const high = Number.parseInt(thresholdHighInput.value, 10);
  if (Number.isNaN(low) || Number.isNaN(mid) || Number.isNaN(high) || low >= mid || mid >= high) {
    thresholdStatus.textContent = "値は 低 < 中 < 高 の順で入力してください。";
    return;
  }
  if (!chrome?.storage?.local) {
    thresholdStatus.textContent = "保存先が利用できません。";
    return;
  }
  chrome.storage.local.set({ [COLOR_THRESHOLDS_KEY]: { low, mid, high } }, () => {
    thresholdStatus.textContent = `保存しました (〜${low}% / 〜${mid}% / 〜${high}% / 超)`;
  });
});

thresholdResetButton.addEventListener("click", () => {
  thresholdLowInput.value = 50;
  thresholdMidInput.value = 70;
  thresholdHighInput.value = 90;
  if (!chrome?.storage?.local) {
    thresholdStatus.textContent = "保存先が利用できません。";
    return;
  }
  chrome.storage.local.remove(COLOR_THRESHOLDS_KEY, () => {
    thresholdStatus.textContent = "リセットしました (50/70/90)";
  });
});

loadContactsCount();
loadExcludeKeywords();
loadExcludeRules();
loadRoomEmailDomain();
loadWorkHours();
loadColorThresholds();
void fetchDebugInfo();
