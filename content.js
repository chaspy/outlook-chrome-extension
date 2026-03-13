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
  const SEARCH_HINTS_ID = "oce-calendar-search-hints";
  const SEARCH_HINTS_LIMIT = 10;
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
    /^辞退:/,
    /^予定のキャンセル:/
  ];
  const IGNORE_STATUS_PATTERNS = [/,\s*Free\b/i, /,\s*空き\b/, /,\s*空き時間\b/];
  const ATTENDEE_PLACEHOLDERS = [
    "Invite attendees",
    "Invite required attendees",
    "出席者を追加",
    "必須出席者を招待します"
  ];
  const TITLE_PLACEHOLDERS = ["Add a title", "Add title", "タイトルを追加", "タイトルの追加"];
  const TITLE_ARIA_LABELS = ["Add details for the event", "タイトル"];
  const IGNORE_CALENDAR_NAMES = new Set([
    "Calendar",
    "Birthdays",
    "Japan holidays",
    "予定表",
    "カレンダー"
  ]);
  const TIME_INSIGHTS_ID = "oce-time-insights";
  const TIME_STORAGE_KEY = "oceMeetingTime";
  const EXCLUDE_KEYWORDS_KEY = "oceMeetingExcludeKeywords";
  const WORK_HOURS_KEY = "oceMeetingWorkHours";
  const DEFAULT_WORK_START = 9;
  const DEFAULT_WORK_END = 18;
  const IGNORE_STATUS_FOR_TIME = [
    /\bFree\b/i, /\b空き\b/, /\b空き時間\b/,
    /\bOut of Office\b/i, /\b外出中\b/
  ];
  const TIME_RANGE_PATTERNS = [
    /(\d{1,2}:\d{2}\s*[AP]M)\s+to\s+(\d{1,2}:\d{2}\s*[AP]M)/i,
    /(\d{1,2}:\d{2})\s*(?:から|[〜~\-–])\s*(\d{1,2}:\d{2})/,
    /(\d{1,2}:\d{2}\s*[AP]M)\s*[〜~\-–]\s*(\d{1,2}:\d{2}\s*[AP]M)/i
  ];
  const DATE_PATTERNS_EN = /\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})\b/i;
  const DATE_PATTERNS_JA = /(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/;
  const KNOWN_STATUSES = new Set([
    "Busy", "Free", "Tentative", "Out of Office", "Working Elsewhere",
    "予定あり", "空き", "空き時間", "仮の予定", "外出中", "他の場所で作業"
  ]);
  const WORK_DAYS_PER_WEEK = 5;
  const MONTH_NAMES = {
    january: 0, february: 1, march: 2, april: 3, may: 4, june: 5,
    july: 6, august: 7, september: 8, october: 9, november: 10, december: 11
  };
  const MAX_WEEKS_STORED = 8;

  const ATTENDEE_AUTOFILL_ATTR = "data-oce-attendees-filled";
  const ATTENDEE_AUTOFILLING_ATTR = "data-oce-attendees-filling";
  const EMAIL_PATTERN = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;
  const RECENT_INPUT_TTL_MS = 20000;
  const CONFLICT_HOST_ID = "oce-conflict-host";
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
    contactsByNameIds: new Map(),
    contactsList: [],
    contactsCount: 0,
    contactsLoaded: false,
    showAllClicked: false,
    showAllAttemptedAt: 0,
    ignoredErrors: 0,
    lastIgnoredError: "",
    meetingTimeData: {},
    meetingExcludeKeywords: [],
    meetingWorkStart: DEFAULT_WORK_START,
    meetingWorkEnd: DEFAULT_WORK_END,
    lastTimeItemIds: ""
  };

  let selectionObserver = null;
  let selectionObserverRoot = null;
  let pendingUpdate = false;
  let pendingSelectionUpdate = false;

  const createCopyButton = (textToCopy, label) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "oce-copy-btn";
    btn.title = `${label}をコピー`;
    btn.innerHTML =
      '<svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor">' +
      '<path d="M4 4v-2a2 2 0 0 1 2-2h6a2 2 0 0 1 2 2v6a2 2 0 0 1-2 2h-2v2a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2zm2-2v2h2a2 2 0 0 1 2 2v2h2V2H6zM2 6v6h6V6H2z"/>' +
      "</svg>";
    btn.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      globalThis.navigator.clipboard.writeText(textToCopy).then(() => {
        showToast(`コピーしました: ${textToCopy}`);
      });
    });
    return btn;
  };

  const showToast = (message) => {
    const existing = document.getElementById(TOAST_ID);
    if (existing) existing.remove();

    const toast = document.createElement("div");
    toast.id = TOAST_ID;
    toast.textContent = message;
    document.body.appendChild(toast);

    globalThis.setTimeout(() => {
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

  const extractTimeRange = (label) => {
    for (const pattern of TIME_RANGE_PATTERNS) {
      const match = label.match(pattern);
      if (match) return { startStr: match[1], endStr: match[2] };
    }
    return null;
  };

  const parseTimeToMinutes = (timeStr) => {
    const trimmed = timeStr.trim();
    const ampmMatch = trimmed.match(/^(\d{1,2}):(\d{2})\s*([AP]M)$/i);
    if (ampmMatch) {
      let hours = Number.parseInt(ampmMatch[1], 10);
      const minutes = Number.parseInt(ampmMatch[2], 10);
      const period = ampmMatch[3].toUpperCase();
      if (period === "PM" && hours !== 12) hours += 12;
      if (period === "AM" && hours === 12) hours = 0;
      return hours * 60 + minutes;
    }
    const h24Match = trimmed.match(/^(\d{1,2}):(\d{2})$/);
    if (h24Match) {
      return Number.parseInt(h24Match[1], 10) * 60 + Number.parseInt(h24Match[2], 10);
    }
    return null;
  };

  const getDurationHours = (startStr, endStr) => {
    const startMin = parseTimeToMinutes(startStr);
    const endMin = parseTimeToMinutes(endStr);
    if (startMin === null || endMin === null) return 0;
    const workStartMin = state.meetingWorkStart * 60;
    const workEndMin = state.meetingWorkEnd * 60;
    const clampedStart = Math.max(startMin, workStartMin);
    const clampedEnd = Math.min(endMin, workEndMin);
    const diff = clampedEnd - clampedStart;
    if (diff <= 0) return 0;
    return diff / 60;
  };

  const extractDate = (label) => {
    const enMatch = label.match(DATE_PATTERNS_EN);
    if (enMatch) {
      const month = MONTH_NAMES[enMatch[1].toLowerCase()];
      if (month !== undefined) {
        return new Date(Number.parseInt(enMatch[3], 10), month, Number.parseInt(enMatch[2], 10));
      }
    }
    const jaMatch = label.match(DATE_PATTERNS_JA);
    if (jaMatch) {
      return new Date(Number.parseInt(jaMatch[1], 10), Number.parseInt(jaMatch[2], 10) - 1, Number.parseInt(jaMatch[3], 10));
    }
    return null;
  };

  const extractStatus = (label) => {
    const segments = label.split(",").map((s) => s.trim());
    for (let i = segments.length - 1; i >= 0; i -= 1) {
      if (KNOWN_STATUSES.has(segments[i])) return segments[i];
    }
    return segments.length > 0 ? segments[segments.length - 1] : "";
  };

  const isAllDayEvent = (el) => !el.closest(".templateColumnContent");

  const isRecurringEvent = (el) => {
    const icons = el.querySelectorAll("i, span.ms-Icon, [class*='icon']");
    for (const icon of icons) {
      const cls = icon.className || "";
      const title = icon.getAttribute("title") || "";
      if (/recur|repeat/i.test(cls) || /recur|repeat|繰り返し|定期/i.test(title)) return true;
    }
    return false;
  };

  const isMeetingForTimeCalc = (el) => {
    if (isAllDayEvent(el)) return false;
    const label = getAriaLabel(el);
    if (!label) return false;
    if (IGNORE_LABEL_PATTERNS.some((p) => p.test(label))) return false;
    const status = extractStatus(label);
    if (IGNORE_STATUS_FOR_TIME.some((p) => p.test(status))) return false;
    if (state.meetingExcludeKeywords.length > 0) {
      const lowerLabel = label.toLowerCase();
      if (state.meetingExcludeKeywords.some((kw) => lowerLabel.includes(kw.toLowerCase()))) return false;
    }
    if (!extractTimeRange(label)) return false;
    return true;
  };

  const getWeekMonday = (date) => {
    const d = new Date(date);
    const day = d.getDay();
    const diff = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + diff);
    d.setHours(0, 0, 0, 0);
    return d;
  };

  const formatWeekKey = (monday) => {
    const y = monday.getFullYear();
    const m = String(monday.getMonth() + 1).padStart(2, "0");
    const d = String(monday.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  };

  const parseEventEntry = (el) => {
    const label = getAriaLabel(el);
    const timeRange = extractTimeRange(label);
    if (!timeRange) return null;
    const date = extractDate(label);
    if (!date) return null;
    const duration = getDurationHours(timeRange.startStr, timeRange.endStr);
    if (duration <= 0) return null;
    return { date, duration, recurring: isRecurringEvent(el) };
  };

  const OOF_STATUS_PATTERNS = [/\bOut of Office\b/i, /\b外出中\b/];

  const isOofStatus = (status) => OOF_STATUS_PATTERNS.some((p) => p.test(status));

  const collectOofDays = () => {
    const allEvents = collectEvents();
    const oofByWeek = {};
    for (const el of allEvents) {
      if (!isAllDayEvent(el)) continue;
      const label = getAriaLabel(el);
      const status = extractStatus(label);
      if (!isOofStatus(status)) continue;
      const date = extractDate(label);
      if (!date) continue;
      const day = date.getDay();
      if (day === 0 || day === 6) continue;
      const weekKey = formatWeekKey(getWeekMonday(date));
      if (!oofByWeek[weekKey]) oofByWeek[weekKey] = new Set();
      oofByWeek[weekKey].add(formatWeekKey(date));
    }
    const result = {};
    for (const [weekKey, dates] of Object.entries(oofByWeek)) {
      result[weekKey] = dates.size;
    }
    return result;
  };

  const addEventToWeek = (weeks, weekKey, duration, recurring) => {
    if (!weeks[weekKey]) weeks[weekKey] = { total: 0, recurring: 0, oneTime: 0, count: 0, oofDays: 0 };
    weeks[weekKey].total += duration;
    weeks[weekKey].count += 1;
    if (recurring) {
      weeks[weekKey].recurring += duration;
    } else {
      weeks[weekKey].oneTime += duration;
    }
  };

  const collectMeetingData = () => {
    const events = collectEvents().filter(isMeetingForTimeCalc);
    const weeks = {};
    const seenIds = new Set();
    for (const el of events) {
      const calItemId = el.dataset.calitemid;
      if (calItemId && seenIds.has(calItemId)) continue;
      if (calItemId) seenIds.add(calItemId);
      const entry = parseEventEntry(el);
      if (!entry) continue;
      const weekKey = formatWeekKey(getWeekMonday(entry.date));
      addEventToWeek(weeks, weekKey, entry.duration, entry.recurring);
    }
    const oofDays = collectOofDays();
    for (const key of Object.keys(weeks)) {
      weeks[key].total = Math.round(weeks[key].total * 10) / 10;
      weeks[key].recurring = Math.round(weeks[key].recurring * 10) / 10;
      weeks[key].oneTime = Math.round(weeks[key].oneTime * 10) / 10;
      weeks[key].oofDays = oofDays[key] || 0;
    }
    return weeks;
  };

  const saveMeetingTimeData = (newWeekData) => {
    const merged = { ...state.meetingTimeData };
    for (const [key, value] of Object.entries(newWeekData)) {
      merged[key] = value;
    }
    const keys = Object.keys(merged).sort((a, b) => a.localeCompare(b)).reverse();
    const pruned = {};
    for (let i = 0; i < Math.min(keys.length, MAX_WEEKS_STORED); i += 1) {
      pruned[keys[i]] = merged[keys[i]];
    }
    state.meetingTimeData = pruned;
    if (chrome?.storage?.local) {
      chrome.storage.local.set({ [TIME_STORAGE_KEY]: pruned });
    }
  };

  const clearMeetingTimeCache = () => {
    state.lastTimeItemIds = "";
  };

  const loadMeetingTimeData = () => {
    if (!chrome?.storage?.local) return;
    chrome.storage.local.get([TIME_STORAGE_KEY, EXCLUDE_KEYWORDS_KEY, WORK_HOURS_KEY], (result) => {
      if (result[TIME_STORAGE_KEY]) {
        state.meetingTimeData = result[TIME_STORAGE_KEY];
      }
      if (Array.isArray(result[EXCLUDE_KEYWORDS_KEY])) {
        state.meetingExcludeKeywords = result[EXCLUDE_KEYWORDS_KEY];
      }
      if (result[WORK_HOURS_KEY]) {
        state.meetingWorkStart = result[WORK_HOURS_KEY].start ?? DEFAULT_WORK_START;
        state.meetingWorkEnd = result[WORK_HOURS_KEY].end ?? DEFAULT_WORK_END;
      }
    });
  };

  const watchTimeSettings = () => {
    if (!chrome?.storage?.onChanged) return;
    chrome.storage.onChanged.addListener((changes, area) => {
      if (area !== "local") return;
      let needsUpdate = false;
      if (changes[EXCLUDE_KEYWORDS_KEY]) {
        const newValue = changes[EXCLUDE_KEYWORDS_KEY].newValue;
        state.meetingExcludeKeywords = Array.isArray(newValue) ? newValue : [];
        needsUpdate = true;
      }
      if (changes[WORK_HOURS_KEY]) {
        const wh = changes[WORK_HOURS_KEY].newValue;
        state.meetingWorkStart = wh?.start ?? DEFAULT_WORK_START;
        state.meetingWorkEnd = wh?.end ?? DEFAULT_WORK_END;
        needsUpdate = true;
      }
      if (needsUpdate) {
        clearMeetingTimeCache();
        scheduleTimeUpdate();
      }
    });
  };

  const getWeekCapacity = (weekData) => {
    const hoursPerDay = state.meetingWorkEnd - state.meetingWorkStart;
    const days = WORK_DAYS_PER_WEEK - (weekData?.oofDays || 0);
    return hoursPerDay * Math.max(days, 1);
  };

  const formatWeekLabel = (weekKey) => {
    const parts = weekKey.split("-");
    return `${Number.parseInt(parts[1], 10)}/${Number.parseInt(parts[2], 10)}`;
  };

  const ensureTimeInsights = () => {
    if (document.getElementById(TIME_INSIGHTS_ID)) return;
    const summary = document.getElementById(SELECTED_SUMMARY_ID);
    if (!summary) return;

    const container = document.createElement("div");
    container.id = TIME_INSIGHTS_ID;
    container.innerHTML = "";

    summary.after(container);
    renderTimeInsights();
  };

  const renderTimeInsights = () => {
    const container = document.getElementById(TIME_INSIGHTS_ID);
    if (!container) return;
    container.textContent = "";

    const data = state.meetingTimeData;
    const keys = Object.keys(data).sort((a, b) => a.localeCompare(b));
    if (keys.length === 0) {
      const empty = document.createElement("div");
      empty.className = "oce-time-header";
      empty.textContent = "ミーティング時間: データなし";
      container.appendChild(empty);
      return;
    }

    const today = new Date();
    const currentMonday = getWeekMonday(today);
    const currentWeekKey = formatWeekKey(currentMonday);
    const currentData = data[currentWeekKey];

    const header = document.createElement("div");
    header.className = "oce-time-header";
    const title = document.createElement("span");
    title.textContent = "ミーティング時間";
    header.appendChild(title);
    if (currentData) {
      const pct = Math.round((currentData.total / getWeekCapacity(currentData)) * 100);
      const current = document.createElement("span");
      current.className = "oce-time-current";
      current.textContent = `今週: ${currentData.total}h (${pct}%)`;
      header.appendChild(current);
    }
    container.appendChild(header);

    const maxTotal = Math.max(...keys.map((k) => data[k].total), 1);

    const bars = document.createElement("div");
    bars.className = "oce-time-bars";
    for (const key of keys) {
      const week = data[key];
      const row = document.createElement("div");
      row.className = "oce-time-bar-row";
      if (key === currentWeekKey) row.dataset.current = "true";

      const label = document.createElement("span");
      label.className = "oce-time-bar-label";
      label.textContent = formatWeekLabel(key);
      row.appendChild(label);

      const track = document.createElement("div");
      track.className = "oce-time-bar-track";

      const recurPct = (week.recurring / maxTotal) * 100;
      const oneTimePct = (week.oneTime / maxTotal) * 100;

      const recurBar = document.createElement("div");
      recurBar.className = "oce-time-bar-recurring";
      recurBar.style.width = `${recurPct}%`;
      track.appendChild(recurBar);

      const oneTimeBar = document.createElement("div");
      oneTimeBar.className = "oce-time-bar-onetime";
      oneTimeBar.style.width = `${oneTimePct}%`;
      track.appendChild(oneTimeBar);

      row.appendChild(track);

      const weekPct = Math.round((week.total / getWeekCapacity(week)) * 100);
      const hours = document.createElement("span");
      hours.className = "oce-time-bar-hours";
      hours.textContent = `${week.total}h (${weekPct}%)`;
      row.appendChild(hours);

      bars.appendChild(row);
    }
    container.appendChild(bars);

    const legend = document.createElement("div");
    legend.className = "oce-time-legend";

    const recurLegend = document.createElement("span");
    recurLegend.className = "oce-time-legend-item";
    const recurSwatch = document.createElement("span");
    recurSwatch.className = "oce-time-legend-swatch recurring";
    recurLegend.appendChild(recurSwatch);
    recurLegend.append("定期");
    legend.appendChild(recurLegend);

    const oneLegend = document.createElement("span");
    oneLegend.className = "oce-time-legend-item";
    const oneSwatch = document.createElement("span");
    oneSwatch.className = "oce-time-legend-swatch onetime";
    oneLegend.appendChild(oneSwatch);
    oneLegend.append("単発");
    legend.appendChild(oneLegend);

    container.appendChild(legend);
  };

  let pendingTimeUpdate = null;

  const updateTimeInsights = () => {
    const events = collectEvents();
    const currentIds = events
      .filter((el) => el.dataset.calitemid)
      .map((el) => el.dataset.calitemid)
      .sort((a, b) => a.localeCompare(b))
      .join(",");
    if (currentIds === state.lastTimeItemIds && currentIds.length > 0) return;
    state.lastTimeItemIds = currentIds;

    const newData = collectMeetingData();
    saveMeetingTimeData(newData);
    if (globalThis.requestAnimationFrame) {
      globalThis.requestAnimationFrame(renderTimeInsights);
    } else {
      renderTimeInsights();
    }
  };

  const scheduleTimeUpdate = () => {
    if (pendingTimeUpdate) {
      globalThis.clearTimeout(pendingTimeUpdate);
    }
    pendingTimeUpdate = globalThis.setTimeout(() => {
      pendingTimeUpdate = null;
      updateTimeInsights();
    }, 2000);
  };

  const searchUtils = globalThis.oceSearchUtils;
  if (!searchUtils) return;
  const {
    stripInvisible,
    normalizeText,
    normalizeNameKey,
    normalizeEmail,
    normalizeId,
    tokenizeSearchTerm,
    matchesTokens
  } = searchUtils;

  const isEffectivelyEmpty = (value) => stripInvisible(value).length === 0;
  const isEmailInput = (value) => value.includes("@");

  const captureIgnoredError = (error, context) => {
    state.ignoredErrors += 1;
    const message = error?.message || String(error || "");
    state.lastIgnoredError = context ? `${context}: ${message}` : message;
  };

  const updateContactsFromList = (list) => {
    const byName = new Map();
    const byEmail = new Map();
    const byNameIds = new Map();
    const uniqueList = new Map();
    let count = 0;
    if (Array.isArray(list)) {
      list.forEach((entry) => {
        if (!entry) return;
        const name = normalizeText(entry.name || "");
        const nameKey = normalizeNameKey(name);
        const email = normalizeEmail(entry.email || "");
        const id = normalizeId(entry.id || "");
        if (!nameKey || !email) return;
        const uniqueKey = `${nameKey}\n${email}\n${id}`;
        if (!uniqueList.has(uniqueKey)) {
          uniqueList.set(uniqueKey, {
            name,
            nameKey,
            email,
            id
          });
        }
        if (!byName.has(nameKey)) byName.set(nameKey, new Set());
        byName.get(nameKey).add(email);
        if (!byEmail.has(email)) byEmail.set(email, new Set());
        byEmail.get(email).add(nameKey);
        if (id) {
          if (!byNameIds.has(nameKey)) byNameIds.set(nameKey, new Set());
          byNameIds.get(nameKey).add(id);
        }
        count += 1;
      });
    }
    state.contactsByName = byName;
    state.contactsByEmail = byEmail;
    state.contactsByNameIds = byNameIds;
    state.contactsList = [...uniqueList.values()];
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
  const getIdsForName = (nameKey) => state.contactsByNameIds.get(nameKey);
  const getContactsList = () => state.contactsList || [];

  const getUniqueEmailForName = (nameKey) => {
    const emails = getEmailsForName(nameKey);
    if (emails?.size !== 1) return "";
    return [...emails][0];
  };

  const sleep = (ms) =>
    new Promise((resolve) => {
      globalThis.setTimeout(resolve, ms);
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
      if (globalThis.parent?.document) docs.add(globalThis.parent.document);
    } catch (error) {
      captureIgnoredError(error, "parent.document");
    }
    try {
      if (globalThis.top?.document) docs.add(globalThis.top.document);
    } catch (error) {
      captureIgnoredError(error, "top.document");
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
        captureIgnoredError(error, "frame.contentDocument");
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

  const matchesCalendarSearch = (name, tokens) => {
    if (!tokens || tokens.length === 0) return true;
    const nameKey = normalizeNameKey(name);
    const emails = getEmailsForName(nameKey);
    const ids = getIdsForName(nameKey);
    return matchesTokens(tokens, { nameKey, emails, ids });
  };

  const findCalendarMatches = (term) => {
    const tokens = tokenizeSearchTerm(term);
    const matches = [];
    const matchedNameKeys = new Set();
    let total = 0;
    if (tokens.length === 0) {
      return { matches, matchedNameKeys, total };
    }
    const buttons = getCalendarOptionButtons();
    const seen = new Set();
    for (const button of buttons) {
      const label = button.querySelector(".ATH58");
      const raw = label ? label.textContent : button.textContent;
      const name = normalizeText(raw || "");
      if (!name) continue;
      const nameKey = normalizeNameKey(name);
      if (seen.has(nameKey)) continue;
      seen.add(nameKey);
      if (!matchesCalendarSearch(name, tokens)) continue;
      matchedNameKeys.add(nameKey);
      total += 1;
      if (matches.length < SEARCH_HINTS_LIMIT) {
        matches.push({ name, nameKey });
      }
    }
    return { matches, matchedNameKeys, total };
  };

  const findContactMatches = (term, excludedNameKeys) => {
    const tokens = tokenizeSearchTerm(term);
    if (tokens.length === 0) return { matches: [], total: 0 };
    const matches = [];
    let total = 0;
    const excluded = excludedNameKeys || new Set();
    for (const entry of getContactsList()) {
      if (excluded.has(entry.nameKey)) continue;
      const emails = entry.email ? [entry.email] : [];
      const ids = entry.id ? [entry.id] : [];
      if (!matchesTokens(tokens, { nameKey: entry.nameKey, emails, ids })) {
        continue;
      }
      total += 1;
      if (matches.length < SEARCH_HINTS_LIMIT) {
        matches.push(entry);
      }
    }
    return { matches, total };
  };

  const setNativeInputValue = (input, value) => {
    const prototype = Object.getPrototypeOf(input);
    const descriptor = Object.getOwnPropertyDescriptor(prototype, "value");
    const setter = descriptor?.set;
    if (setter) {
      setter.call(input, value);
    } else {
      input.value = value;
    }
  };

  const findOutlookSearchInput = () => {
    const docs = getAccessibleDocuments();
    for (const doc of docs) {
      const input = doc.getElementById("topSearchInput");
      if (input) return input;
      const box = doc.getElementById("searchBoxId-Calendar");
      if (box) {
        const candidate = box.querySelector(
          "input[aria-label=\"Search\"], input[type=\"search\"], input"
        );
        if (candidate) return candidate;
      }
    }
    return null;
  };

  const registerContactInOutlook = (entry) => {
    const input = findOutlookSearchInput();
    if (!input) {
      showToast("Outlookの検索欄が見つかりません");
      return;
    }
    const value = entry.email || entry.name;
    setNativeInputValue(input, value);
    input.focus();
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    showToast("Outlookの検索欄に入力しました");
  };

  const SHOW_ALL_LABELS = ["Show all", "すべて表示"];
  const RIBBON_CONTAINER_SELECTOR =
    "#innerRibbonContainer, [data-automation-type=\"RibbonBottomBarContainer\"]";

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
    const ribbonContainer = document.querySelector(RIBBON_CONTAINER_SELECTOR);
    if (!ribbonContainer) return null;
    let host = ribbonContainer.querySelector(`#${CONFLICT_HOST_ID}`);
    if (!host) {
      host = document.createElement("div");
      host.id = CONFLICT_HOST_ID;
      host.className = "oce-conflict-host";
      ribbonContainer.appendChild(host);
    }
    return host;
  };

  const placeConflictButton = (button) => {
    const anchor = findConflictButtonAnchor();
    if (anchor?.isConnected) {
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

  const isCalendarSelected = (nameKey) =>
    findCalendarButtonsForName(nameKey).some(
      (button) => button.getAttribute("aria-selected") === "true"
    );

  const selectCalendar = (nameKey) => {
    const buttons = findCalendarButtonsForName(nameKey);
    const target = buttons.find(
      (button) => button.getAttribute("aria-selected") !== "true"
    );
    if (target) {
      target.click();
      showToast("選択しました");
      return true;
    }
    if (buttons.length > 0) {
      showToast("選択済みです");
      return true;
    }
    return false;
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

      const nameActions = document.createElement("span");
      nameActions.className = "oce-action-group";

      nameActions.appendChild(createCopyButton(name, "名前"));

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
      nameActions.appendChild(removeButton);
      nameRow.appendChild(nameActions);

      pill.appendChild(nameRow);
      if (email) {
        const emailRow = document.createElement("div");
        emailRow.className = "oce-selected-email-row";
        const emailSpan = document.createElement("span");
        emailSpan.className = "oce-selected-email";
        emailSpan.textContent = email;
        emailRow.appendChild(emailSpan);
        emailRow.appendChild(createCopyButton(email, "メールアドレス"));
        pill.appendChild(emailRow);
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
      ensureTimeInsights();
      scheduleTimeUpdate();
    };
    if (globalThis.requestAnimationFrame) {
      globalThis.requestAnimationFrame(run);
    } else {
      globalThis.setTimeout(run, 100);
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
    if (globalThis.requestAnimationFrame) {
      globalThis.requestAnimationFrame(run);
    } else {
      globalThis.setTimeout(run, 100);
    }
  };

  const clearSearchClasses = (groups, rows) => {
    for (const group of groups) {
      group.classList.remove(SEARCH_HIT_CLASS, SEARCH_MISS_CLASS);
    }
    for (const row of rows) {
      row.classList.remove(SEARCH_HIT_CLASS, SEARCH_MISS_CLASS);
      const buttons = row.querySelectorAll("button[role=\"option\"]");
      for (const button of buttons) {
        button.classList.remove(SEARCH_HIT_CLASS, SEARCH_MISS_CLASS);
      }
    }
  };

  const setRowMatchState = (row, match) => {
    row.classList.toggle(SEARCH_HIT_CLASS, match);
    row.classList.toggle(SEARCH_MISS_CLASS, !match);
    const buttons = row.querySelectorAll("button[role=\"option\"]");
    for (const button of buttons) {
      button.classList.toggle(SEARCH_HIT_CLASS, match);
      button.classList.toggle(SEARCH_MISS_CLASS, !match);
    }
  };

  const getCalendarRowName = (row) => {
    const label = row.querySelector(".ATH58");
    const raw = label ? label.textContent : row.textContent;
    return normalizeText(raw || "");
  };

  const applyRowMatch = (row, tokens, groupHasMatch, matchState) => {
    const name = getCalendarRowName(row);
    const match = matchesCalendarSearch(name, tokens);
    setRowMatchState(row, match);
    matchState.candidates += 1;
    if (!match) return;
    matchState.matches += 1;
    const group = row.closest("li[aria-label]");
    if (group) groupHasMatch.set(group, true);
    if (!matchState.firstMatch) matchState.firstMatch = row;
  };

  const updateGroupMatchClasses = (groups, groupHasMatch) => {
    for (const group of groups) {
      const hasRows = group.querySelector("button[role=\"option\"]");
      if (!hasRows) continue;
      const hasMatch = groupHasMatch.get(group) === true;
      group.classList.toggle(SEARCH_HIT_CLASS, hasMatch);
      group.classList.toggle(SEARCH_MISS_CLASS, !hasMatch);
    }
  };

  const applySearchToDoc = (doc, term, tokens) => {
    const listRoot = findCalendarListRootInDoc(doc);
    if (!listRoot) return { candidates: 0, matches: 0, firstMatch: null };
    const listContainer = listRoot.ul;
    const groups = [...listContainer.querySelectorAll("li[aria-label]")];
    const rows = getCalendarRows(listContainer);

    if (!tokens || tokens.length === 0) {
      clearSearchClasses(groups, rows);
      return { candidates: 0, matches: 0, firstMatch: null };
    }

    const groupHasMatch = new Map();
    const matchState = { candidates: 0, matches: 0, firstMatch: null };

    for (const row of rows) {
      applyRowMatch(row, tokens, groupHasMatch, matchState);
    }

    updateGroupMatchClasses(groups, groupHasMatch);

    return matchState;
  };

  const applyCalendarSearch = (raw = "") => {
    const rawTerm = raw;
    const tokens = tokenizeSearchTerm(rawTerm);
    const docs = collectDocuments(document);
    state.searchCandidates = 0;
    state.searchMatches = 0;

    let firstMatch = null;

    for (const doc of docs) {
      const result = applySearchToDoc(doc, rawTerm, tokens);
      state.searchCandidates += result.candidates;
      state.searchMatches += result.matches;
      if (!firstMatch && result.firstMatch) firstMatch = result.firstMatch;
    }

    if (firstMatch && rawTerm !== state.searchTerm) {
      firstMatch.scrollIntoView({ block: "center", inline: "nearest" });
    }
    state.searchTerm = rawTerm;
    state.lastSearchAppliedAt = Date.now();
    renderSelectedSummary();
    renderSearchHints(rawTerm);
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
    input.placeholder = "検索（名前/メール/ID）";
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

    const hints = document.createElement("div");
    hints.id = SEARCH_HINTS_ID;
    hints.className = "oce-search-hints";
    hints.hidden = true;
    container.appendChild(hints);

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

    searchBox.after(summary);
    renderSelectedSummary();
  };

  const createHintNameRow = (name) => {
    const nameRow = document.createElement("div");
    nameRow.className = "oce-search-hints-name-row";
    const nameEl = document.createElement("div");
    nameEl.className = "oce-search-hints-name";
    nameEl.textContent = name;
    nameRow.appendChild(nameEl);
    nameRow.appendChild(createCopyButton(name, "名前"));
    return nameRow;
  };

  const renderSearchHints = (rawTerm) => {
    const hints = document.getElementById(SEARCH_HINTS_ID);
    if (!hints) return;
    hints.textContent = "";
    hints.hidden = true;

    if (!state.contactsLoaded) return;
    const tokens = tokenizeSearchTerm(rawTerm);
    if (tokens.length === 0) return;

    const calendarMatches = findCalendarMatches(rawTerm);
    const contactMatches = findContactMatches(rawTerm, calendarMatches.matchedNameKeys);
    if (calendarMatches.total === 0 && contactMatches.total === 0) return;

    hints.hidden = false;

    const renderSection = (titleText, total, items, renderItem) => {
      const section = document.createElement("div");
      section.className = "oce-search-section";

      const title = document.createElement("div");
      title.className = "oce-search-hints-title";
      title.textContent = titleText;

      const count = document.createElement("span");
      count.className = "oce-search-hints-count";
      count.textContent = `${total}`;
      title.appendChild(count);

      section.appendChild(title);

      if (items.length === 0) {
        const empty = document.createElement("div");
        empty.className = "oce-search-hints-empty";
        empty.textContent = "該当なし";
        section.appendChild(empty);
      } else {
        const list = document.createElement("div");
        list.className = "oce-search-hints-list";
        items.forEach((item) => list.appendChild(renderItem(item)));
        section.appendChild(list);
      }

      if (total > items.length) {
        const more = document.createElement("div");
        more.className = "oce-search-hints-more";
        more.textContent = `他${total - items.length}件`;
        section.appendChild(more);
      }

      hints.appendChild(section);
    };

    renderSection(
      "予定表リストの一致",
      calendarMatches.total,
      calendarMatches.matches,
      (entry) => {
        const row = document.createElement("div");
        row.className = "oce-search-hints-row";

        const info = document.createElement("div");
        info.className = "oce-search-hints-info";

        info.appendChild(createHintNameRow(entry.name));

        const emails = getEmailsForName(entry.nameKey);
        if (emails && emails.size > 0) {
          const emailRow = document.createElement("div");
          emailRow.className = "oce-search-hints-email-row";
          const email = document.createElement("div");
          email.className = "oce-search-hints-email";
          const list = [...emails];
          const preview = list.slice(0, 2).join(", ");
          email.textContent =
            list.length > 2 ? `${preview} 他${list.length - 2}件` : preview;
          emailRow.appendChild(email);
          if (list.length === 1) {
            emailRow.appendChild(createCopyButton(list[0], "メールアドレス"));
          }
          info.appendChild(emailRow);
        }

        const ids = getIdsForName(entry.nameKey);
        if (ids && ids.size > 0) {
          const id = document.createElement("div");
          id.className = "oce-search-hints-id";
          const list = [...ids];
          const preview = list.slice(0, 2).map((value) => `@${value}`).join(", ");
          id.textContent =
            list.length > 2 ? `${preview} 他${list.length - 2}件` : preview;
          info.appendChild(id);
        }

        row.appendChild(info);

        const action = document.createElement("button");
        action.type = "button";
        action.className = "oce-search-hints-action";
        if (isCalendarSelected(entry.nameKey)) {
          action.textContent = "解除";
          action.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            if (!deselectCalendar(entry.nameKey)) {
              showToast("解除できませんでした");
            }
          });
        } else {
          action.textContent = "選択";
          action.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            if (!selectCalendar(entry.nameKey)) {
              showToast("選択できませんでした");
            }
          });
        }
        row.appendChild(action);

        return row;
      }
    );

    renderSection(
      "登録済みのみの一致",
      contactMatches.total,
      contactMatches.matches,
      (entry) => {
        const row = document.createElement("div");
        row.className = "oce-search-hints-row";

        const info = document.createElement("div");
        info.className = "oce-search-hints-info";

        info.appendChild(createHintNameRow(entry.name));

        const emailRow = document.createElement("div");
        emailRow.className = "oce-search-hints-email-row";
        const email = document.createElement("div");
        email.className = "oce-search-hints-email";
        email.textContent = entry.email;
        emailRow.appendChild(email);
        emailRow.appendChild(createCopyButton(entry.email, "メールアドレス"));
        info.appendChild(emailRow);

        if (entry.id) {
          const id = document.createElement("div");
          id.className = "oce-search-hints-id";
          id.textContent = `@${entry.id}`;
          info.appendChild(id);
        }

        row.appendChild(info);

        const action = document.createElement("button");
        action.type = "button";
        action.className = "oce-search-hints-action";
        action.textContent = "登録";
        action.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          registerContactInOutlook(entry);
        });
        row.appendChild(action);

        return row;
      }
    );
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

  const trimEmailToken = (token) => {
    const allowed = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789._%+-@";
    let start = 0;
    let end = token.length;
    while (start < end && !allowed.includes(token[start])) start += 1;
    while (end > start && !allowed.includes(token[end - 1])) end -= 1;
    return token.slice(start, end);
  };

  const isLikelyEmail = (value) => {
    if (!value) return false;
    const at = value.indexOf("@");
    if (at <= 0 || at !== value.lastIndexOf("@")) return false;

    const local = value.slice(0, at);
    const domain = value.slice(at + 1);
    if (!local || !domain) return false;
    if (local.startsWith(".") || local.endsWith(".")) return false;
    if (domain.startsWith(".") || domain.endsWith(".")) return false;
    if (domain.includes("..")) return false;

    const dot = domain.lastIndexOf(".");
    if (dot <= 0) return false;
    const tld = domain.slice(dot + 1);
    return tld.length >= 2;
  };

  const getEditorPillEmailKeys = (editor) => {
    const pills = [...editor.querySelectorAll("._EType_RECIPIENT_ENTITY")];
    const emails = new Set();
    pills.forEach((pill) => {
      const source = `${pill.getAttribute("aria-label") || ""} ${pill.textContent || ""}`;
      const tokens = source.split(/\s+/);
      tokens.forEach((token) => {
        const candidate = trimEmailToken(token);
        if (!isLikelyEmail(candidate)) return;
        const key = normalizeEmail(candidate);
        if (key) emails.add(key);
      });
    });
    return emails;
  };

  const extractFirstEmail = (text) => {
    if (!text) return "";
    const match = text.match(EMAIL_PATTERN);
    return match ? match[0] : "";
  };

  const getSelfEmailCandidates = (root) => {
    if (!root?.querySelectorAll) return [];
    return [
      ...root.querySelectorAll(".mDZXX, .fui-Dropdown__button, .ms-Dropdown-title, .ms-Dropdown")
    ];
  };

  const findSelfEmail = (doc, preferredRoot) => {
    const rootDoc = doc || document;
    const roots = [preferredRoot, rootDoc].filter(Boolean);
    for (const root of roots) {
      const nodes = getSelfEmailCandidates(root);
      for (const node of nodes) {
        const fromAria = extractFirstEmail(node.getAttribute("aria-label") || "");
        if (fromAria) return fromAria;
        const fromText = extractFirstEmail(node.textContent || "");
        if (fromText) return fromText;
      }
    }
    return "";
  };

  const ensureSelfEmail = (doc, preferredRoot) => {
    const preferredEmail = findSelfEmail(doc, preferredRoot);
    if (preferredEmail) {
      state.selfEmail = preferredEmail;
      return preferredEmail;
    }
    if (state.selfEmail) return state.selfEmail;
    const fallbackEmail = findSelfEmail(doc, doc?.body || document.body);
    if (fallbackEmail) state.selfEmail = fallbackEmail;
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
        editor.dataset.placeholder || editor.getAttribute("aria-label") || "";
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

  const getComposeRootForEditor = (editor) =>
    editor.closest("[data-app-section=\"CalendarQuickCompose\"]") ||
    editor.closest("[data-app-section=\"Form_Content\"]") ||
    editor.closest("[role=\"dialog\"]") ||
    editor.ownerDocument?.body ||
    document.body;

  const isTitleInputField = (input) => {
    const placeholder = input.getAttribute("placeholder") || "";
    if (TITLE_PLACEHOLDERS.some((value) => placeholder.includes(value))) return true;

    const ariaLabel = input.getAttribute("aria-label") || "";
    return TITLE_ARIA_LABELS.some((value) => ariaLabel.includes(value));
  };

  const findTitleInputForEditor = (editor) => {
    const root = getComposeRootForEditor(editor);
    const candidates = [...root.querySelectorAll("input")].filter(isVisible);
    return candidates.find((input) => isTitleInputField(input)) || null;
  };

  const isBlankNewEventDraft = (editor) => {
    const titleInput = findTitleInputForEditor(editor);
    if (!titleInput) return false;

    const hasTitle = !isEffectivelyEmpty(titleInput.value || "");
    if (hasTitle) return false;

    const hasPills = editor.querySelectorAll("._EType_RECIPIENT_ENTITY").length > 0;
    if (hasPills) return false;

    const hasAttendeeText = !isEffectivelyEmpty(editor.textContent || "");
    return !hasAttendeeText;
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

  const placeCaretAtEnd = (editor) => {
    editor.focus();
    const selection = globalThis.getSelection?.();
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
      if (!parent?.closest("._EType_RECIPIENT_ENTITY")) {
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
    editor.append(text);
    editor.dispatchEvent(
      new InputEvent("input", { bubbles: true, data: text, inputType: "insertText" })
    );
  };

  const insertText = (editor, text) => {
    clearEditorInputText(editor);
    appendText(editor, text);
  };

  const waitForPillInsert = async (editor, beforeCount, input, timeoutMs = 2000) => {
    const normalizedInput = isEmailInput(input)
      ? normalizeEmail(input)
      : normalizeNameKey(input);
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      const currentCount = editor.querySelectorAll("._EType_RECIPIENT_ENTITY").length;
      if (currentCount > beforeCount) return true;
      if (isEmailInput(input)) {
        const emails = getEditorPillEmailKeys(editor);
        if (emails.has(normalizedInput)) return true;
      } else {
        const labels = getEditorPillLabels(editor);
        if (labels.has(normalizedInput)) return true;
      }
      await sleep(80);
    }
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

    clearEditorInputText(editor);
    await sleep(80);
    return false;
  };

  const resolveAttendeeInputs = (names, selfEmail) => {
    const entries = [];
    const selfEmailKey = normalizeEmail(selfEmail || "");
    const selfNames = selfEmailKey ? getNamesForEmail(selfEmailKey) : null;
    names.forEach((name) => {
      const nameKey = normalizeNameKey(name);
      if (selfNames?.has(nameKey)) return;
      const email = getUniqueEmailForName(nameKey);
      if (email) {
        const emailKey = normalizeEmail(email);
        if (selfEmailKey && emailKey === selfEmailKey) return;
        entries.push({
          input: email,
          nameKey,
          emailKey,
          source: "contact"
        });
      }
    });

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

  const shouldSkipAutofillEditor = (editor) => {
    if (!editor) return true;
    if (editor.getAttribute(ATTENDEE_AUTOFILL_ATTR) === "true") return true;
    if (editor.getAttribute(ATTENDEE_AUTOFILLING_ATTR) === "true") return true;
    if (!state.contactsLoaded && chrome?.storage?.local) return true;
    return !isBlankNewEventDraft(editor);
  };

  const buildAutofillEntries = (editor) => {
    const names = getSelectedCalendarNames();
    const selfEmail = ensureSelfEmail(
      editor.ownerDocument,
      getComposeRootForEditor(editor)
    );
    const entries = resolveAttendeeInputs(names, selfEmail);
    if (entries.length === 0) return null;

    const autofillKey = buildAutofillKey(entries);
    const now = Date.now();
    if (state.lastAutofillKey === autofillKey && now - state.lastAutofillAt < 8000) {
      state.autofillSkips += 1;
      return null;
    }

    state.lastAutofillKey = autofillKey;
    state.lastAutofillAt = now;
    state.autofillRuns += 1;
    state.autofillLastInputs = entries.map((entry) => entry.input).slice(0, 10);
    return entries;
  };

  const getAutofillEntryKey = (entry) => entry.emailKey || entry.nameKey || "";

  const shouldSkipAutofillEntry = (entry, key, existing, existingEmails, seen) => {
    if (!key || wasRecentlyInserted(key) || seen.has(key)) return true;
    if (entry.nameKey && existing.has(entry.nameKey)) return true;
    if (entry.emailKey && existingEmails.has(entry.emailKey)) return true;
    if (entry.emailKey && hasExistingForEmail(entry.emailKey, existing)) return true;
    return false;
  };

  const applyAutofillEntry = async (editor, entry, key, existing, existingEmails) => {
    const inserted = await addAttendeeByName(editor, entry.input);
    if (!inserted) return;
    markInserted(key);
    if (entry.nameKey) existing.add(entry.nameKey);
    if (entry.emailKey) {
      addExistingForEmail(entry.emailKey, existing);
      existingEmails.add(entry.emailKey);
    }
  };

  const applyAutofillEntries = async (editor, entries, existing, existingEmails) => {
    const seen = new Set();
    for (const entry of entries) {
      const key = getAutofillEntryKey(entry);
      if (shouldSkipAutofillEntry(entry, key, existing, existingEmails, seen)) continue;
      seen.add(key);
      await applyAutofillEntry(editor, entry, key, existing, existingEmails);
      await sleep(120);
    }
  };

  const fillAttendees = async (editor) => {
    if (shouldSkipAutofillEditor(editor)) return;
    const entries = buildAutofillEntries(editor);
    if (!entries) return;

    editor.setAttribute(ATTENDEE_AUTOFILL_ATTR, "true");
    editor.setAttribute(ATTENDEE_AUTOFILLING_ATTR, "true");
    const existing = getEditorPillLabels(editor);
    const existingEmails = getEditorPillEmailKeys(editor);

    try {
      await applyAutofillEntries(editor, entries, existing, existingEmails);
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
        const value = (node.dataset.placeholder || "").trim();
        if (value) placeholderSamples.push(value);
      });
    });

    const editors = getAttendeeEditors();
    const editorSummaries = editors.slice(0, 6).map((editor) => ({
      placeholder: editor.dataset.placeholder || null,
      ariaLabel: editor.getAttribute("aria-label"),
      className: editor.className,
      textSample: (editor.textContent || "").trim().slice(0, 120),
      ownerLocation: editor.ownerDocument?.location?.href || ""
    }));

    const searchInput = document.getElementById(SEARCH_INPUT_ID);
    const searchInputValue = searchInput ? searchInput.value : "";
    const searchBoxPresent = !!document.getElementById(SEARCH_BOX_ID);
    const effectiveSearchRaw = searchInputValue || state.searchTerm || "";
    const effectiveSearchTokens = tokenizeSearchTerm(effectiveSearchRaw);

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
          match: matchesCalendarSearch(name, effectiveSearchTokens),
          hasHitClass: row.classList.contains(SEARCH_HIT_CLASS),
          hasMissClass: row.classList.contains(SEARCH_MISS_CLASS)
        });
      }
    });

    const contactSamples = [];
    for (const [name, emails] of state.contactsByName.entries()) {
      if (contactSamples.length >= 10) break;
      const ids = state.contactsByNameIds.get(name);
      contactSamples.push({
        nameKey: name,
        emails: [...emails],
        ids: ids ? [...ids] : []
      });
    }

    const iframes = [...document.querySelectorAll("iframe")].map((frame) => {
      let sameOrigin = false;
      let href = "";
      try {
        sameOrigin = !!frame.contentDocument;
        href = frame.contentDocument?.location?.href || "";
      } catch (error) {
        captureIgnoredError(error, "iframe.contentDocument");
      }
      return { src: frame.src, sameOrigin, href };
    });

    return {
      timestamp: new Date().toISOString(),
      location: globalThis.location?.href || "",
      topLevel: globalThis.top === globalThis.parent,
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
      contactsByNameIdsCount: state.contactsByNameIds.size,
      contactsLoaded: state.contactsLoaded,
      contactsSamples: contactSamples,
      ignoredErrors: state.ignoredErrors,
      lastIgnoredError: state.lastIgnoredError,
      searchTerm: state.searchTerm,
      searchInputValue,
      searchTokens: effectiveSearchTokens,
      searchBoxPresent,
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
      userAgent: navigator.userAgent,
      timeInsights: (() => {
        const allEvents = collectEvents();
        const meetingEvents = allEvents.filter(isMeetingForTimeCalc);
        const counted = meetingEvents.map((el) => {
          const label = getAriaLabel(el);
          const timeRange = extractTimeRange(label);
          const date = extractDate(label);
          const duration = timeRange ? getDurationHours(timeRange.startStr, timeRange.endStr) : 0;
          return {
            title: label.split("、")[0].slice(0, 80),
            time: timeRange ? `${timeRange.startStr}-${timeRange.endStr}` : "",
            date: date ? formatWeekKey(date) : "",
            duration,
            status: extractStatus(label),
            recurring: isRecurringEvent(el)
          };
        });
        return {
          totalEvents: allEvents.length,
          meetingEvents: meetingEvents.length,
          storedData: state.meetingTimeData,
          countedEvents: counted
        };
      })()
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
        if (el.dataset.conflict === "1") {
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
    loadMeetingTimeData();
    watchTimeSettings();
    ensureButton();
    ensureSearchBox();
    ensureSelectedSummary();
    ensureTimeInsights();
    ensureSelectionObserver();
    ensureShowAllExpanded();
    maybeAutofillAttendees();
    scheduleTimeUpdate();
    startObserver();
  };

  if (typeof chrome !== "undefined" && chrome.runtime?.onMessage) {
    chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
      if (message?.type !== "OCE_DEBUG") return;
      try {
        sendResponse(collectDebugInfo());
      } catch (error) {
        sendResponse({ error: error?.message || String(error) });
      }
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot, { once: true });
  } else {
    boot();
  }
})();
