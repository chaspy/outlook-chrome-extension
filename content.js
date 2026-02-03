(() => {
  const BUTTON_ID = "oce-conflict-button";
  const TOAST_ID = "oce-conflict-toast";
  const CONFLICT_CLASS = "oce-conflict";
  const SEARCH_BOX_ID = "oce-calendar-search";
  const SEARCH_INPUT_ID = "oce-calendar-search-input";
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
    searchCandidates: 0,
    searchMatches: 0
  };

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
  const normalizeEmail = (value) => value.trim().toLowerCase();
  const isEmailInput = (value) => value.includes("@");

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
        const key = normalizeText(name);
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
    return buttons.filter((button) => button.querySelector(".ATH58"));
  };

  const findCalendarListRoot = () => {
    const docs = collectDocuments(document);
    for (const doc of docs) {
      const list = [...doc.querySelectorAll("ul")].find((ul) =>
        ul.querySelector("button[role=\"option\"] .ATH58")
      );
      if (list) return list;
    }
    return null;
  };

  const applyCalendarSearch = (raw) => {
    const term = normalizeText(raw || "").toLowerCase();
    const docs = collectDocuments(document);
    state.searchCandidates = 0;
    state.searchMatches = 0;

    let firstMatch = null;

    const getRows = (root) =>
      [...root.querySelectorAll(".GsziR")].filter((row) =>
        row.querySelector("button[role=\"option\"] .ATH58")
      );

    docs.forEach((doc) => {
      const groups = [...doc.querySelectorAll("li[aria-label]")];
      const rows = getRows(doc);

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
        const name = normalizeText(label ? label.textContent || "" : "");
        const match = name.toLowerCase().includes(term);
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
        const hasRows = group.querySelector(".GsziR");
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
    input.placeholder = "検索（名前）";
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

  const getEditorPillLabels = (editor) => {
    const pills = [...editor.querySelectorAll("._EType_RECIPIENT_ENTITY")];
    return new Set(
      pills
        .map((pill) => {
          const label = pill.getAttribute("aria-label");
          if (label) return normalizeText(label);
          const text = pill.querySelector(".textContainer-390");
          return text ? normalizeText(text.textContent || "") : "";
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
      : normalizeText(input);
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

  const buildAutofillKey = (inputs) => {
    const normalized = inputs
      .map((input) =>
        isEmailInput(input) ? normalizeEmail(input) : normalizeText(input)
      )
      .filter(Boolean)
      .sort();
    return normalized.join("|");
  };

  const fillAttendees = async (editor) => {
    if (!editor || editor.getAttribute(ATTENDEE_AUTOFILL_ATTR) === "true") return;
    if (editor.getAttribute(ATTENDEE_AUTOFILLING_ATTR) === "true") return;
    const hasPills =
      editor.querySelectorAll("._EType_RECIPIENT_ENTITY").length > 0;
    const hasText = !isEffectivelyEmpty(editor.textContent || "");
    if (hasText && !hasPills) return;

    const names = getSelectedCalendarNames();
    const selfEmail = ensureSelfEmail(editor.ownerDocument);
    const inputs = [...names];
    if (selfEmail) inputs.push(selfEmail);
    if (inputs.length === 0) return;

    const autofillKey = buildAutofillKey(inputs);
    const now = Date.now();
    if (state.lastAutofillKey === autofillKey && now - state.lastAutofillAt < 8000) {
      state.autofillSkips += 1;
      return;
    }

    state.lastAutofillKey = autofillKey;
    state.lastAutofillAt = now;
    state.autofillRuns += 1;
    state.autofillLastInputs = inputs.slice(0, 10);
    editor.setAttribute(ATTENDEE_AUTOFILL_ATTR, "true");
    editor.setAttribute(ATTENDEE_AUTOFILLING_ATTR, "true");

    const existing = getEditorPillLabels(editor);
    const seen = new Set();

    try {
      for (const input of inputs) {
        const key = isEmailInput(input) ? normalizeEmail(input) : normalizeText(input);
        if (wasRecentlyInserted(key)) continue;
        if (seen.has(key)) continue;
        seen.add(key);
        if (!isEmailInput(input) && existing.has(key)) continue;
        const inserted = await addAttendeeByName(editor, input);
        if (inserted) {
          markInserted(key);
          if (!isEmailInput(input)) existing.add(key);
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
      searchTerm: state.searchTerm,
      searchCandidates: state.searchCandidates,
      searchMatches: state.searchMatches,
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
    if (document.getElementById(BUTTON_ID)) return;
    if (!document.querySelector(CALENDAR_ROOT_SELECTOR)) return;

    const button = document.createElement("button");
    button.id = BUTTON_ID;
    button.type = "button";
    button.textContent = "重複検出";
    button.dataset.active = "false";

    button.addEventListener("click", () => {
      toggleDetection();
    });

    document.body.appendChild(button);
  };

  const startObserver = () => {
    const observer = new MutationObserver(() => {
      ensureButton();
      ensureSearchBox();
      if (state.searchTerm) applyCalendarSearch(state.searchTerm);
      maybeAutofillAttendees();
    });

    observer.observe(document.documentElement, {
      childList: true,
      subtree: true
    });
  };

  const boot = () => {
    ensureButton();
    ensureSearchBox();
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
