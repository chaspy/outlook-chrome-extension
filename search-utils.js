(() => {
  const stripInvisible = (value) =>
    value
      .replaceAll(/\s/g, "")
      .replaceAll("\u200B", "")
      .replaceAll("\u200C", "")
      .replaceAll("\u200D", "")
      .replaceAll("\uFEFF", "");

  const stripZeroWidth = (value) =>
    value
      .replaceAll("\u200B", "")
      .replaceAll("\u200C", "")
      .replaceAll("\u200D", "")
      .replaceAll("\uFEFF", "");

  const normalizeText = (value) => stripInvisible(value).replaceAll(/\s+/g, " ").trim();

  const normalizeNameKey = (value) =>
    stripInvisible(value).replaceAll(/\s+/g, "").trim().toLowerCase();

  const normalizeEmail = (value) => value.trim().toLowerCase();

  const normalizeId = (value) =>
    normalizeText(value)
      .replaceAll(/\s+/g, "")
      .replace(/^@+/, "")
      .toLowerCase();

  const tokenizeSearchTerm = (value) => {
    const raw = stripZeroWidth(value || "");
    if (!raw) return [];
    const cleaned = raw
      .replaceAll("\u3000", " ")
      .replaceAll(/[<>"',;]+/g, " ")
      .toLowerCase();
    return cleaned
      .split(/\s+/)
      .map((token) => token.trim())
      .filter((token) => token.length > 0 && token !== "@");
  };

  const isLooseMatch = (haystack, needle) => {
    if (!haystack || !needle) return false;
    return haystack.includes(needle);
  };

  const matchesToken = (token, { nameKey, emails, ids }) => {
    const termKey = normalizeNameKey(token);
    if (termKey && isLooseMatch(nameKey, termKey)) return true;

    const emailTerm = normalizeEmail(token);
    for (const email of emails || []) {
      if (isLooseMatch(normalizeEmail(email), emailTerm)) return true;
    }

    const idTerm = normalizeId(token);
    for (const id of ids || []) {
      if (isLooseMatch(normalizeId(id), idTerm)) return true;
    }

    return false;
  };

  const matchesTokens = (tokens, entry) => tokens.every((token) => matchesToken(token, entry));

  const api = {
    stripInvisible,
    stripZeroWidth,
    normalizeText,
    normalizeNameKey,
    normalizeEmail,
    normalizeId,
    tokenizeSearchTerm,
    isLooseMatch,
    matchesToken,
    matchesTokens
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  } else {
    globalThis.oceSearchUtils = api;
  }
})();
