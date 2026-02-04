(() => {
  const BUTTON_ID = "oce-conflict-button";
  const TOAST_ID = "oce-conflict-toast";
  const CONFLICT_CLASS = "oce-conflict";
  const SEARCH_BOX_ID = "oce-calendar-search";
  const SEARCH_INPUT_ID = "oce-calendar-search-input";
  const SELECTED_SUMMARY_ID = "oce-calendar-selected-summary";
  const SELECTED_COUNT_ID = "oce-calendar-selected-count";
  const SELECTED_LIST_ID = "oce-calendar-selected-list";
  const SEARCH_HIT_CLASS = "oce-calendar-search-hit";
  const SEARCH_MISS_CLASS = "oce-calendar-search-miss";
  const CALENDAR_ROOT_SELECTOR = ".templateColumnContent, [data-calitemid]";
  const USE_OUTLOOK_CONFLICT_FLAG = false;
  const IGNORE_LABEL_PATTERNS = [
    /^Canceled:/i,
    /^Canceled event\b/i,
    /^Cancelled:/i,
    /^Cancelled event\b/i,
    /^Declined:/i,
    /^キャンセル済み:/,
    /^キャンセルイベント/,
    /^辞退:/
  ];
  const IGNORE_STATUS_PATTERNS = [/,\s*Free\b/i, /,\s*空き\b/, /,\s*空き時間\b/];
  const ATTENDEE_PLACEHOLDERS = ["Invite attendees", "出席者を追加"];
  const IGNORE_CALENDAR_NAMES = new Set(["Calendar", "Birthdays", "Japan holidays"]);
  const ATTENDEE_AUTOFILL_ATTR = "data-oce-attendees-filled";
  const ATTENDEE_AUTOFILLING_ATTR = "data-oce-attendees-filling";
  const EMAIL_PATTERN = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;
  const RECENT_INPUT_TTL_MS = 20000;
  const CONTACTS_STORAGE_KEY = "oceContacts";

  const state = {
    active: false,
    lastRunAt: 0,
    selfEmail: "",
    lastAutofillKey: "",
    lastAutofillAt: 0,
    autofillRuns: 0,
    autofillSkips: 0,
    autofillLastInputs: [],
    recentInputs: new Map(),
    searchTerm: "",
    lastSearchAppliedAt: 0,
    searchCandidates: 0,
    searchMatches: 0,
    contactsByName: new Map(),
    contactsByEmail: new Map(),
    contactsCount: 0,
    contactsLoaded: false,
    showAllClicked: false,
    showAllAttemptedAt: 0
  };

  let selectionObserver = null;
  let selectionObserverRoot = null;
  let pendingUpdate = false;
  let pendingSelectionUpdate = false;

  const showToast = (message) => {
    const existing = document.getElementById(TOAST_ID);
    if (existing) existing.remove();

    const toast = document.createElement("div");
    toast.id = TOAST_ID;
    toast.textContent = message;
    document.body.appendChild(toast);

    window.setTimeout(() => {
      toast.remove();
    }, 2200);
  };

  const clearHighlights = () => {
    document.querySelectorAll(`.${CONFLICT_CLASS}`).forEach((el) => {
      el.classList.remove(CONFLICT_CLASS);
    });
  };

  const getAriaLabel = (el) => {
    const direct = el.getAttribute("aria-label");
    if (direct) return direct;
    const child = el.querySelector("[aria-label]");
    return child ? child.getAttribute("aria-label") || "" : "";
  };

  const isIgnorable = (el) => {
    const label = getAriaLabel(el);
    return (
      IGNORE_LABEL_PATTERNS.some((pattern) => pattern.test(label)) ||
      IGNORE_STATUS_PATTERNS.some((pattern) => pattern.test(label))
    );
  };

  const stripInvisible = (value) =>
    value.replace(/[\s\u200B\u200C\u200D\uFEFF]/g, "");
  const isEffectivelyEmpty = (value) => stripInvisible(value).length === 0;
  const normalizeText = (value) =>
    stripInvisible(value).replace(/\s+/g, " ").trim();
  const normalizeNameKey = (value) =>
    stripInvisible(value).replace(/\s+/g, "").trim().toLowerCase();
  const normalizeEmail = (value) => value.trim().toLowerCase();
  const isEmailInput = (value) => value.includes("@");

  const updateContactsFromList = (list) => {
    const byName = new Map();
    const byEmail = new Map();
    let count = 0;
    if (Array.isArray(list)) {
      list.forEach((entry) => {
        if (!entry) return;
        const name = normalizeText(entry.name || "");
        const nameKey = normalizeNameKey(name);
        const email = normalizeEmail(entry.email || "");
        if (!nameKey || !email) return;
        if (!byName.has(nameKey)) byName.set(nameKey, new Set());
        byName.get(nameKey).add(email);
        if (!byEmail.has(email)) byEmail.set(email, new Set());
        byEmail.get(email).add(nameKey);
        count += 1;
      });
    }
    state.contactsByName = byName;
    state.contactsByEmail = byEmail;
    state.contactsCount = count;
    state.contactsLoaded = true;
  };

  const loadContacts = () => {
    if (!chrome?.storage?.local) {
      state.contactsLoaded = true;
      return;
    }
    chrome.storage.local.get(CONTACTS_STORAGE_KEY, (result) => {
      updateContactsFromList(result[CONTACTS_STORAGE_KEY]);
      if (state.searchTerm) applyCalendarSearch(state.searchTerm);
      renderSelectedSummary();
      maybeAutofillAttendees();
    });
  };

  const watchContacts = () => {
    if (!chrome?.storage?.onChanged) return;
    chrome.storage.onChanged.addListener((changes, area) => {
      if (area !== "local") return;
      if (!changes[CONTACTS_STORAGE_KEY]) return;
      updateContactsFromList(changes[CONTACTS_STORAGE_KEY].newValue);
      if (state.searchTerm) applyCalendarSearch(state.searchTerm);
      renderSelectedSummary();
      maybeAutofillAttendees();
    });
  };

  const getEmailsForName = (nameKey) => state.contactsByName.get(nameKey);
  const getNamesForEmail = (emailKey) => state.contactsByEmail.get(emailKey);

  const getUniqueEmailForName = (nameKey) => {
    const emails = getEmailsForName(nameKey);
    if (!emails || emails.size !== 1) return "";
    return [...emails][0];
  };

  const sleep = (ms) =>
    new Promise((resolve) => {
      window.setTimeout(resolve, ms);
    });

  const pruneRecentInputs = () => {
    const now = Date.now();
    [...state.recentInputs.entries()].forEach(([key, timestamp]) => {
      if (now - timestamp > RECENT_INPUT_TTL_MS) state.recentInputs.delete(key);
    });
  };

  const wasRecentlyInserted = (key) => {
    pruneRecentInputs();
    return state.recentInputs.has(key);
  };

  const markInserted = (key) => {
    state.recentInputs.set(key, Date.now());
  };

  const isVisible = (el) =>
    !!(el && (el.offsetWidth || el.offsetHeight || el.getClientRects().length));

  const getAccessibleDocuments = () => {
    const docs = new Set();
    docs.add(document);
    try {
      if (window.parent?.document) docs.add(window.parent.document);
    } catch (error) {
      // ignore cross-origin frames
    }
    try {
      if (window.top?.document) docs.add(window.top.document);
    } catch (error) {
      // ignore cross-origin frames
    }
    return [...docs];
  };

  const collectDocuments = (rootDoc, depth = 0, maxDepth = 2) => {
    const docs = [rootDoc];
    if (depth >= maxDepth) return docs;
    const frames = [...rootDoc.querySelectorAll("iframe")];
    frames.forEach((frame) => {
      try {
        if (frame.contentDocument) {
          docs.push(...collectDocuments(frame.contentDocument, depth + 1, maxDepth));
        }
      } catch (error) {
        // ignore cross-origin frames
      }
    });
    return docs;
  };

  const getSelectedCalendarNames = () => {
    const docs = getAccessibleDocuments();
    const selected = [];
    docs.forEach((doc) => {
      selected.push(
        ...doc.querySelectorAll("button[role=\"option\"][aria-selected=\"true\"]")
      );
    });
    const seen = new Set();
    return selected
      .map((button) => {
        const label = button.querySelector(".ATH58");
        const raw = label ? label.textContent : button.textContent;
        return raw ? normalizeText(raw) : "";
      })
      .filter((name) => name && !IGNORE_CALENDAR_NAMES.has(name))
      .filter((name) => {
        const key = normalizeNameKey(name);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  };

  const getCalendarOptionButtons = () => {
    const docs = collectDocuments(document);
    const buttons = [];
    docs.forEach((doc) => {
      buttons.push(...doc.querySelectorAll("button[role=\"option\"]"));
    });
    return buttons;
  };

  const findCalendarListRootInDoc = (doc) => {
    const lists = [...doc.querySelectorAll("ul")]
      .map((ul) => ({
        ul,
        count: ul.querySelectorAll("button[role=\"option\"]").length
      }))
      .filter((entry) => entry.count >= 5)
      .sort((a, b) => b.count - a.count);
    return lists.length > 0 ? lists[0] : null;
  };

  const findCalendarListRoot = () => {
    const docs = collectDocuments(document);
    let best = null;
    docs.forEach((doc) => {
      const candidate = findCalendarListRootInDoc(doc);
      if (!candidate) return;
      if (!best || candidate.count > best.count) best = candidate;
    });
    return best ? best.ul : null;
  };

  

  const getCalendarRows = (root) => {
    const buttons = [...root.querySelectorAll("button[role=\"option\"]")].filter(
      (button) =>
        !button.closest(
          ".ms-FloatingSuggestions, .ms-Suggestions, .ms-BasePicker, .ms-BaseFloatingPicker"
        )
    );
    const rows = new Set();
    buttons.forEach((button) => {
      const row =
        button.closest(".GsziR") ||
        button.closest("div[draggable]") ||
        button.closest("li") ||
        button;
      rows.add(row);
    });
    return [...rows];
  };

  const matchesCalendarSearch = (name, term) => {
    if (!term) return true;
    const nameKey = normalizeNameKey(name);
    const termKey = normalizeNameKey(term);
    if (termKey && nameKey.includes(termKey)) return true;
    const emails = getEmailsForName(nameKey);
    if (!emails) return false;
    for (const email of emails) {
      if (normalizeEmail(email).includes(term)) return true;
    }
    return false;
  };

  const SHOW_ALL_LABELS = ["Show all", "すべて表示"];
  const TABLIST_SELECTOR = "#tablist, [role=\"tablist\"]";

  const isShowAllLabel = (text) => {
    const normalized = normalizeText(text || "").toLowerCase();
    if (!normalized) return false;
    return SHOW_ALL_LABELS.some((label) =>
      normalized.includes(normalizeText(label).toLowerCase())
    );
  };

  const isVisibleButton = (button) =>
    !!button && button.getClientRects().length > 0 && button.offsetParent !== null;

  const findShowAllButtons = (doc) => {
    const listRoot = findCalendarListRootInDoc(doc);
    const searchRoot = listRoot?.ul?.parentElement || listRoot?.ul || doc;
    return [...searchRoot.querySelectorAll("button")].filter((button) =>
      isShowAllLabel(button.textContent)
    );
  };

  const ensureShowAllExpanded = () => {
    if (state.showAllClicked) return;
    const now = Date.now();
    if (now - state.showAllAttemptedAt < 800) return;
    state.showAllAttemptedAt = now;

    const docs = collectDocuments(document);
    let clicked = false;
    docs.forEach((doc) => {
      findShowAllButtons(doc).forEach((button) => {
        if (button.disabled) return;
        if (button.getAttribute("aria-disabled") === "true") return;
        if (!isVisibleButton(button)) return;
        button.click();
        clicked = true;
      });
    });

    if (clicked) {
      state.showAllClicked = true;
    }
  };

  const findConflictButtonAnchor = () => {
    const tablist = document.querySelector("#tablist") || document.querySelector("[role=\"tablist\"]");
    if (!tablist) return null;
    return tablist.parentElement || tablist;
  };

  const placeConflictButton = (button) => {
    const anchor = findConflictButtonAnchor();
    if (anchor && anchor.isConnected) {
      const computed = window.getComputedStyle(anchor);
      if (computed.position === "static") {
        anchor.style.position = "relative";
      }
      button.classList.add("oce-conflict-inline");
      if (button.parentElement !== anchor) {
        anchor.appendChild(button);
      }
      return;
    }
    button.classList.remove("oce-conflict-inline");
    if (!button.isConnected) document.body.appendChild(button);
  };

  const getSelectedSummaryEntries = () => {
    const names = getSelectedCalendarNames();
    const entries = names.map((name) => {
      const nameKey = normalizeNameKey(name);
      const email = getUniqueEmailForName(nameKey);
      return { name, nameKey, email };
    });
    return { entries, total: names.length };
  };

  const findCalendarButtonsForName = (nameKey) => {
    const buttons = getCalendarOptionButtons();
    return buttons.filter((button) => {
      const label = button.querySelector(".ATH58");
      const raw = label ? label.textContent : button.textContent;
      const key = normalizeNameKey(raw || "");
      return key === nameKey;
    });
  };

  const deselectCalendar = (nameKey) => {
    const buttons = findCalendarButtonsForName(nameKey);
    const target = buttons.find(
      (button) => button.getAttribute("aria-selected") === "true"
    );
    if (target) {
      target.click();
      showToast("選択を解除しました");
      return true;
    }
    return false;
  };

  const renderSelectedSummary = () => {
    const summary = document.getElementById(SELECTED_SUMMARY_ID);
    if (!summary) return;
    const countEl = summary.querySelector(`#${SELECTED_COUNT_ID}`);
    const listEl = summary.querySelector(`#${SELECTED_LIST_ID}`);
    if (!listEl) return;

    const { entries, total } = getSelectedSummaryEntries();
    if (countEl) {
      countEl.textContent = `${total}`;
    }

    listEl.textContent = "";
    if (entries.length === 0) {
      const empty = document.createElement("div");
      empty.className = "oce-selected-empty";
      empty.textContent = "選択なし";
      listEl.appendChild(empty);
      return;
    }

    entries.forEach(({ name, nameKey, email }) => {
      const pill = document.createElement("div");
      pill.className = "oce-selected-pill";

      const nameRow = document.createElement("div");
      nameRow.className = "oce-selected-row";

      const nameText = document.createElement("span");
      nameText.textContent = name;
      nameRow.appendChild(nameText);

      const removeButton = document.createElement("button");
      removeButton.type = "button";
      removeButton.className = "oce-selected-remove";
      removeButton.textContent = "解除";
      removeButton.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        if (!deselectCalendar(nameKey)) {
          showToast("解除できませんでした");
        }
      });
      nameRow.appendChild(removeButton);

      pill.appendChild(nameRow);
      if (email) {
        const emailSpan = document.createElement("span");
        emailSpan.className = "oce-selected-email";
        emailSpan.textContent = email;
        pill.appendChild(emailSpan);
      }
      listEl.appendChild(pill);
    });
  };

  const scheduleUpdate = () => {
    if (pendingUpdate) return;
    pendingUpdate = true;
    const run = () => {
      pendingUpdate = false;
      ensureButton();
      ensureSearchBox();
      ensureSelectedSummary();
      ensureSelectionObserver();
      ensureShowAllExpanded();
      if (state.searchTerm) {
        const now = Date.now();
        if (now - state.lastSearchAppliedAt > 250) {
          applyCalendarSearch(state.searchTerm);
        }
      }
      maybeAutofillAttendees();
    };
    if (window.requestAnimationFrame) {
      window.requestAnimationFrame(run);
    } else {
      window.setTimeout(run, 100);
    }
  };

  const scheduleSelectionUpdate = () => {
    if (pendingSelectionUpdate) return;
    pendingSelectionUpdate = true;
    const run = () => {
      pendingSelectionUpdate = false;
      renderSelectedSummary();
      maybeAutofillAttendees();
    };
    if (window.requestAnimationFrame) {
      window.requestAnimationFrame(run);
    } else {
      window.setTimeout(run, 100);
    }
  };


  const applyCalendarSearch = (raw) => {
    const term = normalizeText(raw || "").toLowerCase();
    const docs = collectDocuments(document);
    state.searchCandidates = 0;
    state.searchMatches = 0;

    let firstMatch = null;

    docs.forEach((doc) => {
      const listRoot = findCalendarListRootInDoc(doc);
      if (!listRoot) return;
      const listContainer = listRoot.ul;
      const groups = [...listContainer.querySelectorAll("li[aria-label]")];
      const rows = getCalendarRows(listContainer);

      if (!term) {
        groups.forEach((group) => {
          group.classList.remove(SEARCH_HIT_CLASS, SEARCH_MISS_CLASS);
        });
        rows.forEach((row) => {
          row.classList.remove(SEARCH_HIT_CLASS, SEARCH_MISS_CLASS);
          row
            .querySelectorAll("button[role=\"option\"]")
            .forEach((button) =>
              button.classList.remove(SEARCH_HIT_CLASS, SEARCH_MISS_CLASS)
            );
        });
        return;
      }

      const groupHasMatch = new Map();
      rows.forEach((row) => {
        const label = row.querySelector(".ATH58");
        const raw = label ? label.textContent : row.textContent;
        const name = normalizeText(raw || "");
        const match = matchesCalendarSearch(name, term);
        row.classList.toggle(SEARCH_HIT_CLASS, match);
        row.classList.toggle(SEARCH_MISS_CLASS, !match);
        row
          .querySelectorAll("button[role=\"option\"]")
          .forEach((button) => {
            button.classList.toggle(SEARCH_HIT_CLASS, match);
            button.classList.toggle(SEARCH_MISS_CLASS, !match);
          });
        state.searchCandidates += 1;
        if (match) {
          state.searchMatches += 1;
          const group = row.closest("li[aria-label]");
          if (group) groupHasMatch.set(group, true);
          if (!firstMatch) firstMatch = row;
        }
      });

      groups.forEach((group) => {
        const hasRows = group.querySelector("button[role=\"option\"]");
        if (!hasRows) return;
        const hasMatch = groupHasMatch.get(group) === true;
        group.classList.toggle(SEARCH_HIT_CLASS, hasMatch);
        group.classList.toggle(SEARCH_MISS_CLASS, !hasMatch);
      });
    });

    if (firstMatch && term !== state.searchTerm) {
      firstMatch.scrollIntoView({ block: "center", inline: "nearest" });
    }
    state.searchTerm = term;
    state.lastSearchAppliedAt = Date.now();
    renderSelectedSummary();
  };

  const ensureSearchBox = () => {
    if (document.getElementById(SEARCH_BOX_ID)) return;
    const listRoot = findCalendarListRoot();
    if (!listRoot) return;

    const container = document.createElement("div");
    container.id = SEARCH_BOX_ID;

    const input = document.createElement("input");
    input.id = SEARCH_INPUT_ID;
    input.type = "search";
    input.placeholder = "検索（名前/メール）";
    input.autocomplete = "off";
    input.spellcheck = false;
    input.value = state.searchTerm;

    input.addEventListener("input", (event) => {
      applyCalendarSearch(event.target.value);
    });
    input.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        input.value = "";
        applyCalendarSearch("");
      }
    });

    container.appendChild(input);
    listRoot.parentElement?.insertBefore(container, listRoot);
    if (state.searchTerm) applyCalendarSearch(state.searchTerm);
  };

  const ensureSelectedSummary = () => {
    if (document.getElementById(SELECTED_SUMMARY_ID)) return;
    const searchBox = document.getElementById(SEARCH_BOX_ID);
    if (!searchBox) return;

    const summary = document.createElement("div");
    summary.id = SELECTED_SUMMARY_ID;

    const header = document.createElement("div");
    header.className = "oce-selected-header";

    const title = document.createElement("div");
    title.className = "oce-selected-title";
    const titleLabel = document.createElement("span");
    titleLabel.textContent = "選択中";

    const count = document.createElement("span");
    count.id = SELECTED_COUNT_ID;
    count.textContent = "0";
    title.appendChild(titleLabel);
    title.appendChild(count);

    header.appendChild(title);

    const list = document.createElement("div");
    list.id = SELECTED_LIST_ID;
    list.className = "oce-selected-list";

    summary.appendChild(header);
    summary.appendChild(list);

    searchBox.insertAdjacentElement("afterend", summary);
    renderSelectedSummary();
  };

  const ensureSelectionObserver = () => {
    const listRoot = findCalendarListRoot();
    if (!listRoot || listRoot === selectionObserverRoot) return;

    if (selectionObserver) selectionObserver.disconnect();
    selectionObserverRoot = listRoot;
    selectionObserver = new MutationObserver((mutations) => {
      const hasSelectionChange = mutations.some(
        (mutation) =>
          mutation.type === "attributes" && mutation.attributeName === "aria-selected"
      );
      if (hasSelectionChange) {
        scheduleSelectionUpdate();
      }
    });

    selectionObserver.observe(listRoot, {
      attributes: true,
      subtree: true,
      attributeFilter: ["aria-selected"]
    });
  };

  const getEditorPillLabels = (editor) => {
    const pills = [...editor.querySelectorAll("._EType_RECIPIENT_ENTITY")];
    return new Set(
      pills
        .map((pill) => {
          const label = pill.getAttribute("aria-label");
          if (label) return normalizeNameKey(label);
          const text = pill.querySelector(".textContainer-390");
          return text ? normalizeNameKey(text.textContent || "") : "";
        })
        .filter(Boolean)
    );
  };

  const findSelfEmail = (doc) => {
    const root = doc || document;
    const nodes = [...root.querySelectorAll(".ms-Dropdown-title, .ms-Dropdown, [role=\"combobox\"]")];
    for (const node of nodes) {
      const text = node.textContent || "";
      const match = text.match(EMAIL_PATTERN);
      if (match) return match[0];
    }
    return "";
  };

  const ensureSelfEmail = (doc) => {
    if (state.selfEmail) return state.selfEmail;
    const email = findSelfEmail(doc);
    if (email) state.selfEmail = email;
    return state.selfEmail;
  };

  const getAttendeeEditors = () => {
    const docs = collectDocuments(document);
    const editors = [];
    docs.forEach((doc) => {
      editors.push(...doc.querySelectorAll("[contenteditable=\"true\"]"));
    });
    const visibleEditors = editors.filter(isVisible);

    const byPlaceholder = visibleEditors.filter((editor) => {
      const placeholder =
        editor.getAttribute("data-placeholder") || editor.getAttribute("aria-label") || "";
      return ATTENDEE_PLACEHOLDERS.some((value) => placeholder.includes(value));
    });
    if (byPlaceholder.length > 0) return byPlaceholder;

    const pickerParents = [
      ...document.querySelectorAll(".ms-BasePicker, .ms-BaseFloatingPicker")
    ]
      .map((node) => node.parentElement)
      .filter(Boolean);

    const fallback = [];
    pickerParents.forEach((parent) => {
      const editor = parent.querySelector("[contenteditable=\"true\"]");
      if (editor && isVisible(editor)) fallback.push(editor);
    });

    return fallback;
  };

  const findSuggestionItems = (editor) => {
    const selector =
      "[role=\"option\"][aria-label], [data-automationid=\"suggestionItem\"], .ms-Suggestions-item";
    const scopedRoot = editor ? editor.closest(".CoqO5") : null;
    const scoped = scopedRoot ? [...scopedRoot.querySelectorAll(selector)] : [];
    const global = [...document.querySelectorAll(selector)];
    const merged = [...new Set([...scoped, ...global])];
    return merged.filter(isVisible);
  };

  const extractSuggestionName = (item) => {
    const aria = item.getAttribute("aria-label") || "";
    if (aria) {
      const parts = aria.split(/\s[-–]\s/);
      if (parts.length > 0) return normalizeText(parts[0]);
    }

    const personaPrimary =
      item.querySelector(".ms-Persona-primaryText") ||
      item.querySelector("[data-automationid=\"PersonaPrimaryText\"]");
    if (personaPrimary) return normalizeText(personaPrimary.textContent || "");

    const firstSpan = item.querySelector("span");
    if (firstSpan) return normalizeText(firstSpan.textContent || "");

    return normalizeText(item.textContent || "");
  };

  const extractEmailFromSuggestion = (item) => {
    const aria = item.getAttribute("aria-label") || "";
    if (aria) {
      const parts = aria.split(/\s[-–]\s/);
      if (parts.length > 1) {
        const candidate = parts.slice(1).join(" - ");
        if (candidate.includes("@")) return normalizeText(candidate);
      }
    }

    const emailSpan = [...item.querySelectorAll("span")].find((span) =>
      (span.textContent || "").includes("@")
    );
    if (emailSpan) return normalizeText(emailSpan.textContent || "");

    const text = item.textContent || "";
    const match = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,}/i);
    return match ? match[0] : "";
  };

  const findExactSuggestionMatch = (name, editor) => {
    const normalizedName = normalizeText(name);
    const normalizedEmail = normalizeEmail(name);
    const items = findSuggestionItems(editor);
    const matches = items.filter((item) => {
      if (isEmailInput(name)) {
        return normalizeEmail(extractEmailFromSuggestion(item)) === normalizedEmail;
      }
      return extractSuggestionName(item) === normalizedName;
    });
    if (matches.length === 1) return matches[0];
    return null;
  };

  const waitForExactSuggestion = async (name, editor, timeoutMs = 2000) => {
    await sleep(isEmailInput(name) ? 320 : 160);
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      const match = findExactSuggestionMatch(name, editor);
      if (match) return match;
      await sleep(80);
    }
    return null;
  };

  const clearEditor = (editor) => {
    editor.textContent = "";
    editor.dispatchEvent(new InputEvent("input", { bubbles: true }));
  };

  const placeCaretAtEnd = (editor) => {
    editor.focus();
    const selection = window.getSelection();
    if (!selection) return;
    const range = document.createRange();
    range.selectNodeContents(editor);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
  };

  const clearEditorInputText = (editor) => {
    const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
    const toRemove = [];
    let node = walker.nextNode();
    while (node) {
      const parent = node.parentElement;
      if (!parent || !parent.closest("._EType_RECIPIENT_ENTITY")) {
        toRemove.push(node);
      }
      node = walker.nextNode();
    }
    toRemove.forEach((textNode) => textNode.remove());
    editor.dispatchEvent(new InputEvent("input", { bubbles: true }));
  };

  const appendText = (editor, text) => {
    placeCaretAtEnd(editor);
    editor.dispatchEvent(
      new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        data: text,
        inputType: "insertText"
      })
    );
    if (document.queryCommandSupported && document.queryCommandSupported("insertText")) {
      document.execCommand("insertText", false, text);
    } else {
      editor.textContent += text;
    }
    editor.dispatchEvent(
      new InputEvent("input", { bubbles: true, data: text, inputType: "insertText" })
    );
  };

  const emitKey = (editor, key, code, keyCode) => {
    const eventInit = {
      key,
      code,
      keyCode,
      which: keyCode,
      charCode: key.length === 1 ? key.charCodeAt(0) : 0,
      bubbles: true,
      cancelable: true
    };
    editor.dispatchEvent(new KeyboardEvent("keydown", eventInit));
    editor.dispatchEvent(new KeyboardEvent("keypress", eventInit));
    editor.dispatchEvent(new KeyboardEvent("keyup", eventInit));
  };

  const commitByArrowEnter = async (editor) => {
    emitKey(editor, "ArrowDown", "ArrowDown", 40);
    await sleep(120);
    emitKey(editor, "Enter", "Enter", 13);
    editor.dispatchEvent(new Event("change", { bubbles: true }));
  };

  const typeChar = (editor, char, code, keyCode) => {
    emitKey(editor, char, code, keyCode);
    appendText(editor, char);
  };

  const insertText = (editor, text) => {
    clearEditorInputText(editor);
    appendText(editor, text);
  };

  const commitEditor = (editor) => {
    const events = [
      { key: "Enter", code: "Enter", keyCode: 13, which: 13 },
      { key: "Tab", code: "Tab", keyCode: 9, which: 9 }
    ];
    events.forEach((eventInit) => {
      editor.dispatchEvent(
        new KeyboardEvent("keydown", { ...eventInit, bubbles: true, cancelable: true })
      );
      editor.dispatchEvent(
        new KeyboardEvent("keyup", { ...eventInit, bubbles: true, cancelable: true })
      );
    });
    editor.dispatchEvent(new Event("change", { bubbles: true }));
  };

  const waitForPillIncrease = async (editor, beforeCount, timeoutMs = 1500) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      const currentCount = editor.querySelectorAll("._EType_RECIPIENT_ENTITY").length;
      if (currentCount > beforeCount) return true;
      await sleep(80);
    }
    return false;
  };

  const waitForPillInsert = async (editor, beforeCount, input, timeoutMs = 2000) => {
    const normalizedInput = isEmailInput(input)
      ? normalizeEmail(input)
      : normalizeNameKey(input);
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      const currentCount = editor.querySelectorAll("._EType_RECIPIENT_ENTITY").length;
      if (currentCount > beforeCount) return true;
      if (!isEmailInput(input)) {
        const labels = getEditorPillLabels(editor);
        if (labels.has(normalizedInput)) return true;
      }
      await sleep(80);
    }
    return false;
  };

  const attemptDirectEmailCommit = async (editor, email, beforeCount) => {
    clearEditorInputText(editor);
    appendText(editor, email);
    await sleep(80);
    typeChar(editor, ";", "Semicolon", 186);
    await sleep(180);
    commitEditor(editor);
    editor.blur();
    editor.dispatchEvent(new Event("focusout", { bubbles: true }));
    const inserted = await waitForPillInsert(editor, beforeCount, email, 2500);
    if (inserted) {
      clearEditorInputText(editor);
      return true;
    }
    // Leave the text so the user can manually resolve it if needed.
    return false;
  };

  const addAttendeeByName = async (editor, name) => {
    const beforeCount = editor.querySelectorAll("._EType_RECIPIENT_ENTITY").length;
    insertText(editor, name);
    const match = await waitForExactSuggestion(name, editor, isEmailInput(name) ? 3000 : 2000);
    if (match) {
      match.click();
      const inserted = await waitForPillInsert(editor, beforeCount, name, 2500);
      clearEditorInputText(editor);
      return inserted;
    }

    if (isEmailInput(name)) {
      const inserted = await attemptDirectEmailCommit(editor, name, beforeCount);
      if (inserted) return true;
      await commitByArrowEnter(editor);
      return await waitForPillInsert(editor, beforeCount, name, 2500);
    }

    clearEditor(editor);
    await sleep(80);
    return false;
  };

  const resolveAttendeeInputs = (names, selfEmail) => {
    const entries = [];
    names.forEach((name) => {
      const nameKey = normalizeNameKey(name);
      const email = getUniqueEmailForName(nameKey);
      if (email) {
        entries.push({
          input: email,
          nameKey,
          emailKey: normalizeEmail(email),
          source: "contact"
        });
      } else {
        entries.push({ input: name, nameKey, emailKey: "", source: "name" });
      }
    });

    if (selfEmail) {
      const emailKey = normalizeEmail(selfEmail);
      const namesForEmail = getNamesForEmail(emailKey);
      const nameKey =
        namesForEmail && namesForEmail.size === 1 ? [...namesForEmail][0] : "";
      entries.push({ input: selfEmail, nameKey, emailKey, source: "self" });
    }

    return entries;
  };

  const hasExistingForEmail = (emailKey, existingNames) => {
    if (!emailKey) return false;
    const names = getNamesForEmail(emailKey);
    if (!names) return false;
    for (const name of names) {
      if (existingNames.has(name)) return true;
    }
    return false;
  };

  const addExistingForEmail = (emailKey, existingNames) => {
    const names = getNamesForEmail(emailKey);
    if (!names) return;
    names.forEach((name) => existingNames.add(name));
  };

  const buildAutofillKey = (entries) => {
    const normalized = entries
      .map((entry) => entry.emailKey || entry.nameKey || "")
      .filter(Boolean)
      .sort();
    return normalized.join("|");
  };

  const fillAttendees = async (editor) => {
    if (!editor || editor.getAttribute(ATTENDEE_AUTOFILL_ATTR) === "true") return;
    if (editor.getAttribute(ATTENDEE_AUTOFILLING_ATTR) === "true") return;
    if (!state.contactsLoaded && chrome?.storage?.local) return;
    const hasPills =
      editor.querySelectorAll("._EType_RECIPIENT_ENTITY").length > 0;
    const hasText = !isEffectivelyEmpty(editor.textContent || "");
    if (hasText && !hasPills) return;

    const names = getSelectedCalendarNames();
    const selfEmail = ensureSelfEmail(editor.ownerDocument);
    const entries = resolveAttendeeInputs(names, selfEmail);
    if (entries.length === 0) return;

    const autofillKey = buildAutofillKey(entries);
    const now = Date.now();
    if (state.lastAutofillKey === autofillKey && now - state.lastAutofillAt < 8000) {
      state.autofillSkips += 1;
      return;
    }

    state.lastAutofillKey = autofillKey;
    state.lastAutofillAt = now;
    state.autofillRuns += 1;
    state.autofillLastInputs = entries.map((entry) => entry.input).slice(0, 10);
    editor.setAttribute(ATTENDEE_AUTOFILL_ATTR, "true");
    editor.setAttribute(ATTENDEE_AUTOFILLING_ATTR, "true");

    const existing = getEditorPillLabels(editor);
    const seen = new Set();

    try {
      for (const entry of entries) {
        const key = entry.emailKey || entry.nameKey;
        if (wasRecentlyInserted(key)) continue;
        if (seen.has(key)) continue;
        seen.add(key);
        if (entry.nameKey && existing.has(entry.nameKey)) continue;
        if (entry.emailKey && hasExistingForEmail(entry.emailKey, existing)) continue;
        const inserted = await addAttendeeByName(editor, entry.input);
        if (inserted) {
          markInserted(key);
          if (entry.nameKey) existing.add(entry.nameKey);
          if (entry.emailKey) addExistingForEmail(entry.emailKey, existing);
        }
        await sleep(120);
      }
    } finally {
      editor.removeAttribute(ATTENDEE_AUTOFILLING_ATTR);
    }
  };

  const maybeAutofillAttendees = () => {
    const editors = getAttendeeEditors();
    editors.forEach((editor) => {
      if (editor.getAttribute(ATTENDEE_AUTOFILL_ATTR) === "true") return;
      void fillAttendees(editor);
    });
  };

  const collectDebugInfo = () => {
    const docs = collectDocuments(document);
    const contentEditableTotal = docs.reduce(
      (total, doc) => total + doc.querySelectorAll("[contenteditable=\"true\"]").length,
      0
    );

    const placeholderSamples = [];
    docs.forEach((doc) => {
      doc.querySelectorAll("[data-placeholder]").forEach((node) => {
        const value = (node.getAttribute("data-placeholder") || "").trim();
        if (value) placeholderSamples.push(value);
      });
    });

    const editors = getAttendeeEditors();
    const editorSummaries = editors.slice(0, 6).map((editor) => ({
      placeholder: editor.getAttribute("data-placeholder"),
      ariaLabel: editor.getAttribute("aria-label"),
      className: editor.className,
      textSample: (editor.textContent || "").trim().slice(0, 120),
      ownerLocation: editor.ownerDocument?.location?.href || ""
    }));

    const searchInput = document.getElementById(SEARCH_INPUT_ID);
    const searchInputValue = searchInput ? searchInput.value : "";
    const searchBoxPresent = !!document.getElementById(SEARCH_BOX_ID);
    const effectiveSearchTerm = normalizeText(searchInputValue || state.searchTerm || "");

    const rowDiagnostics = [];
    let rowsTotal = 0;
    let optionTotal = 0;
    const listRootDiagnostics = [];
    docs.forEach((doc) => {
      optionTotal += doc.querySelectorAll("button[role=\"option\"]").length;
      const listRoot = findCalendarListRootInDoc(doc);
      if (listRoot) {
        listRootDiagnostics.push({
          optionCount: listRoot.count,
          location: doc.location?.href || ""
        });
      }
      const rows = getCalendarRows(listRoot ? listRoot.ul : doc);
      rowsTotal += rows.length;
      for (const row of rows) {
        if (rowDiagnostics.length >= 20) break;
        const label = row.querySelector(".ATH58");
        const raw = label ? label.textContent : row.textContent;
        const name = normalizeText(raw || "");
        const nameKey = normalizeNameKey(name);
        const emails = getEmailsForName(nameKey)
          ? [...getEmailsForName(nameKey)]
          : [];
        rowDiagnostics.push({
          name,
          nameKey,
          emails,
          match: matchesCalendarSearch(name, effectiveSearchTerm.toLowerCase()),
          hasHitClass: row.classList.contains(SEARCH_HIT_CLASS),
          hasMissClass: row.classList.contains(SEARCH_MISS_CLASS)
        });
      }
    });

    const contactSamples = [];
    for (const [name, emails] of state.contactsByName.entries()) {
      if (contactSamples.length >= 10) break;
      contactSamples.push({ nameKey: name, emails: [...emails] });
    }

    const iframes = [...document.querySelectorAll("iframe")].map((frame) => {
      let sameOrigin = false;
      let href = "";
      try {
        sameOrigin = !!frame.contentDocument;
        href = frame.contentDocument?.location?.href || "";
      } catch (error) {
        // ignore cross-origin frames
      }
      return { src: frame.src, sameOrigin, href };
    });

    return {
      timestamp: new Date().toISOString(),
      location: window.location.href,
      topLevel: window.top === window,
      readyState: document.readyState,
      selfEmail: ensureSelfEmail(document),
      autofillRuns: state.autofillRuns,
      autofillSkips: state.autofillSkips,
      lastAutofillKey: state.lastAutofillKey,
      lastAutofillAt: state.lastAutofillAt,
      recentInputsCount: state.recentInputs.size,
      recentInputsSample: [...state.recentInputs.keys()].slice(0, 10),
      lastAutofillInputs: state.autofillLastInputs,
      contactsCount: state.contactsCount,
      contactsByNameCount: state.contactsByName.size,
      contactsByEmailCount: state.contactsByEmail.size,
      contactsLoaded: state.contactsLoaded,
      contactsSamples: contactSamples,
      searchTerm: state.searchTerm,
      searchInputValue,
      searchBoxPresent,
      effectiveSearchTerm,
      searchCandidates: state.searchCandidates,
      searchMatches: state.searchMatches,
      searchHitClassCount: document.querySelectorAll(`.${SEARCH_HIT_CLASS}`).length,
      searchMissClassCount: document.querySelectorAll(`.${SEARCH_MISS_CLASS}`).length,
      searchOptionTotal: optionTotal,
      searchRowsTotal: rowsTotal,
      searchListRoots: listRootDiagnostics,
      searchRowDiagnostics: rowDiagnostics,
      selectedNames: getSelectedCalendarNames(),
      editorCount: editors.length,
      editorSummaries,
      contentEditableTotal,
      placeholderSamples: [...new Set(placeholderSamples)].slice(0, 20),
      suggestionCount: findSuggestionItems().length,
      iframeCount: iframes.length,
      iframes,
      userAgent: navigator.userAgent
    };
  };

  const groupByContainer = (items) => {
    const groups = new Map();
    items.forEach((item) => {
      const container = item.closest(".templateColumnContent") || document.body;
      if (!groups.has(container)) groups.set(container, []);
      groups.get(container).push(item);
    });
    return [...groups.entries()].map(([container, groupItems]) => ({
      container,
      items: groupItems
    }));
  };

  const detectConflictsInGroup = (container, items) => {
    const parentRect = container.getBoundingClientRect();
    const prepared = items
      .map((el) => {
        const rect = el.getBoundingClientRect();
        return {
          el,
          top: rect.top - parentRect.top,
          bottom: rect.bottom - parentRect.top,
          height: rect.height
        };
      })
      .filter((entry) => entry.height > 2);

    const conflicts = new Set();

    for (let i = 0; i < prepared.length; i += 1) {
      for (let j = i + 1; j < prepared.length; j += 1) {
        const a = prepared[i];
        const b = prepared[j];
        const overlaps = a.top < b.bottom && b.top < a.bottom;
        if (overlaps) {
          conflicts.add(a.el);
          conflicts.add(b.el);
        }
      }
    }

    return conflicts;
  };

  const collectEvents = () => {
    const items = [...document.querySelectorAll("[data-calitemid]")];
    return items.filter((el) => el instanceof HTMLElement);
  };

  const runDetection = () => {
    clearHighlights();

    const events = collectEvents().filter((el) => !isIgnorable(el));
    if (events.length === 0) {
      showToast("予定が見つかりませんでした");
      return;
    }

    const groups = groupByContainer(events);
    const conflicts = new Set();

    groups.forEach(({ container, items }) => {
      const groupConflicts = detectConflictsInGroup(container, items);
      groupConflicts.forEach((el) => conflicts.add(el));
    });

    if (USE_OUTLOOK_CONFLICT_FLAG) {
      events.forEach((el) => {
        if (el.getAttribute("data-conflict") === "1") {
          conflicts.add(el);
        }
      });
    }

    conflicts.forEach((el) => {
      el.classList.add(CONFLICT_CLASS);
    });

    if (conflicts.size === 0) {
      showToast("重複は見つかりませんでした");
      return;
    }

    showToast(`重複候補: ${conflicts.size}件`);
  };

  const toggleDetection = () => {
    state.active = !state.active;
    const button = document.getElementById(BUTTON_ID);
    if (button) button.dataset.active = state.active ? "true" : "false";

    if (state.active) {
      state.lastRunAt = Date.now();
      runDetection();
      if (button) button.textContent = "重複をクリア";
    } else {
      clearHighlights();
      if (button) button.textContent = "重複検出";
      showToast("ハイライトを解除しました");
    }
  };

  const ensureButton = () => {
    const existing = document.getElementById(BUTTON_ID);
    if (existing) {
      placeConflictButton(existing);
      return;
    }
    if (!document.querySelector(CALENDAR_ROOT_SELECTOR)) return;

    const button = document.createElement("button");
    button.id = BUTTON_ID;
    button.type = "button";
    button.textContent = "重複検出";
    button.dataset.active = "false";

    button.addEventListener("click", () => {
      toggleDetection();
    });

    placeConflictButton(button);
  };

  const startObserver = () => {
    const observer = new MutationObserver(() => {
      scheduleUpdate();
    });

    observer.observe(document.documentElement, {
      childList: true,
      subtree: true
    });
  };

  const boot = () => {
    loadContacts();
    watchContacts();
    ensureButton();
    ensureSearchBox();
    ensureSelectedSummary();
    ensureSelectionObserver();
    ensureShowAllExpanded();
    maybeAutofillAttendees();
    startObserver();
  };

  if (typeof chrome !== "undefined" && chrome.runtime?.onMessage) {
    chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
      if (!message || message.type !== "OCE_DEBUG") return false;
      try {
        sendResponse(collectDebugInfo());
      } catch (error) {
        sendResponse({ error: error?.message || String(error) });
      }
      return false;
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot, { once: true });
  } else {
    boot();
  }
})();
