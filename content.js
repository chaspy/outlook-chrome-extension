(() => {
  const BUTTON_ID = "oce-conflict-button";
  const TOAST_ID = "oce-conflict-toast";
  const CONFLICT_CLASS = "oce-conflict";
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

  const state = {
    active: false,
    lastRunAt: 0
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

  const normalizeText = (value) => value.replace(/\s+/g, " ").trim();

  const sleep = (ms) =>
    new Promise((resolve) => {
      window.setTimeout(resolve, ms);
    });

  const isVisible = (el) =>
    !!(el && (el.offsetWidth || el.offsetHeight || el.getClientRects().length));

  const getSelectedCalendarNames = () => {
    const selected = [...document.querySelectorAll("button[role=\"option\"][aria-selected=\"true\"]")];
    return selected
      .map((button) => {
        const label = button.querySelector(".ATH58");
        const raw = label ? label.textContent : button.textContent;
        return raw ? normalizeText(raw) : "";
      })
      .filter((name) => name && !IGNORE_CALENDAR_NAMES.has(name));
  };

  const getAttendeeEditors = () => {
    const editors = [
      ...document.querySelectorAll("[data-placeholder][contenteditable=\"true\"]")
    ];
    return editors.filter((editor) => {
      const placeholder = editor.getAttribute("data-placeholder") || "";
      return ATTENDEE_PLACEHOLDERS.includes(placeholder);
    });
  };

  const findSuggestionItems = () => {
    const items = [
      ...document.querySelectorAll(
        "[role=\"option\"][aria-label], [data-automationid=\"suggestionItem\"], .ms-Suggestions-item"
      )
    ];
    return items.filter(isVisible);
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

  const findExactSuggestionMatch = (name) => {
    const normalizedName = normalizeText(name);
    const items = findSuggestionItems();
    for (const item of items) {
      if (extractSuggestionName(item) === normalizedName) {
        return item;
      }
    }
    return null;
  };

  const waitForExactSuggestion = async (name, timeoutMs = 1600) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      const match = findExactSuggestionMatch(name);
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

  const insertText = (editor, text) => {
    clearEditorInputText(editor);
    placeCaretAtEnd(editor);
    if (document.queryCommandSupported && document.queryCommandSupported("insertText")) {
      document.execCommand("insertText", false, text);
    } else {
      editor.textContent += text;
    }
    editor.dispatchEvent(new InputEvent("input", { bubbles: true }));
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

  const addAttendeeByName = async (editor, name) => {
    insertText(editor, name);
    const match = await waitForExactSuggestion(name);
    if (match) {
      match.click();
      const start = Date.now();
      while (Date.now() - start < 1500) {
        const pill = editor.querySelector("._EType_RECIPIENT_ENTITY[aria-label]");
        if (pill && normalizeText(pill.getAttribute("aria-label") || "") === name) {
          clearEditorInputText(editor);
          return true;
        }
        await sleep(80);
      }
      clearEditorInputText(editor);
      return false;
    }
    clearEditor(editor);
    await sleep(80);
    return false;
  };

  const fillAttendees = async (editor) => {
    if (!editor || editor.getAttribute(ATTENDEE_AUTOFILL_ATTR) === "true") return;
    if (normalizeText(editor.textContent || "") !== "") return;

    const names = getSelectedCalendarNames();
    if (names.length === 0) return;

    editor.setAttribute(ATTENDEE_AUTOFILL_ATTR, "true");

    const seen = new Set();
    for (const name of names) {
      if (seen.has(name)) continue;
      seen.add(name);
      await addAttendeeByName(editor, name);
      await sleep(120);
    }
  };

  const maybeAutofillAttendees = () => {
    const editors = getAttendeeEditors();
    editors.forEach((editor) => {
      if (editor.getAttribute(ATTENDEE_AUTOFILL_ATTR) === "true") return;
      if (normalizeText(editor.textContent || "") !== "") {
        editor.setAttribute(ATTENDEE_AUTOFILL_ATTR, "true");
        return;
      }
      void fillAttendees(editor);
    });
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
      maybeAutofillAttendees();
    });

    observer.observe(document.documentElement, {
      childList: true,
      subtree: true
    });
  };

  const boot = () => {
    ensureButton();
    maybeAutofillAttendees();
    startObserver();
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot, { once: true });
  } else {
    boot();
  }
})();
