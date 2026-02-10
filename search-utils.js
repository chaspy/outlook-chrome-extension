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

  const countMismatches = (left, right) => {
    let mismatches = 0;
    let firstMismatch = -1;
    for (let i = 0; i < left.length; i += 1) {
      if (left[i] !== right[i]) {
        mismatches += 1;
        if (firstMismatch < 0) firstMismatch = i;
        if (mismatches > 2) break;
      }
    }
    return { mismatches, firstMismatch };
  };

  const isSingleSwap = (left, right, index) =>
    index >= 0 &&
    index + 1 < left.length &&
    left[index] === right[index + 1] &&
    left[index + 1] === right[index];

  const isSameLengthDistanceOneOrLess = (left, right) => {
    const { mismatches, firstMismatch } = countMismatches(left, right);
    if (mismatches <= 1) return true;
    if (mismatches === 2) return isSingleSwap(left, right, firstMismatch);
    return false;
  };

  const isInsertDeleteDistanceOneOrLess = (shorter, longer) => {
    let i = 0;
    let j = 0;
    let skips = 0;
    while (i < shorter.length && j < longer.length) {
      if (shorter[i] === longer[j]) {
        i += 1;
        j += 1;
        continue;
      }
      skips += 1;
      if (skips > 1) return false;
      j += 1;
    }
    return true;
  };

  const isEditDistanceOneOrLess = (left, right) => {
    if (left === right) return true;
    const lengthDiff = Math.abs(left.length - right.length);
    if (lengthDiff > 1) return false;
    if (left.length === right.length) {
      return isSameLengthDistanceOneOrLess(left, right);
    }
    const [shorter, longer] = left.length < right.length ? [left, right] : [right, left];
    return isInsertDeleteDistanceOneOrLess(shorter, longer);
  };

  const isLooseMatch = (haystack, needle) => {
    if (!haystack || !needle) return false;
    if (haystack.includes(needle)) return true;
    if (needle.length < 4) return false;
    const minLength = Math.max(1, needle.length - 1);
    const maxLength = needle.length + 1;
    for (let length = minLength; length <= maxLength; length += 1) {
      if (haystack.length < length) continue;
      for (let i = 0; i <= haystack.length - length; i += 1) {
        const slice = haystack.slice(i, i + length);
        if (isEditDistanceOneOrLess(slice, needle)) return true;
      }
    }
    return false;
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
    isEditDistanceOneOrLess,
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
