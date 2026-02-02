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
    });

    observer.observe(document.documentElement, {
      childList: true,
      subtree: true
    });
  };

  const boot = () => {
    ensureButton();
    startObserver();
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot, { once: true });
  } else {
    boot();
  }
})();
