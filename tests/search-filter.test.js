const assert = require("assert/strict");

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
const isEditDistanceOneOrLess = (left, right) => {
  if (left === right) return true;
  const lengthDiff = Math.abs(left.length - right.length);
  if (lengthDiff > 1) return false;

  if (left.length === right.length) {
    let mismatches = 0;
    let firstMismatch = -1;
    for (let i = 0; i < left.length; i += 1) {
      if (left[i] !== right[i]) {
        mismatches += 1;
        if (firstMismatch < 0) firstMismatch = i;
        if (mismatches > 2) return false;
      }
    }
    if (mismatches <= 1) return true;
    if (mismatches === 2 && firstMismatch >= 0 && firstMismatch + 1 < left.length) {
      return (
        left[firstMismatch] === right[firstMismatch + 1] &&
        left[firstMismatch + 1] === right[firstMismatch]
      );
    }
    return false;
  }

  const [shorter, longer] = left.length < right.length ? [left, right] : [right, left];
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
const matchesToken = (token, nameKey, emails, ids) => {
  const termKey = normalizeNameKey(token);
  if (termKey && isLooseMatch(nameKey, termKey)) return true;
  const emailTerm = normalizeEmail(token);
  if (emails) {
    for (const email of emails) {
      if (isLooseMatch(normalizeEmail(email), emailTerm)) return true;
    }
  }
  const idTerm = normalizeId(token);
  if (ids) {
    for (const id of ids) {
      if (isLooseMatch(normalizeId(id), idTerm)) return true;
    }
  }
  return false;
};
const matchesTokens = (tokens, nameKey, emails, ids) =>
  tokens.length === 0 || tokens.every((token) => matchesToken(token, nameKey, emails, ids));

const entry = {
  name: "吉岡 良治",
  email: "ryoji_yoshioka@r.recruit.co.jp",
  id: "ojiry"
};
const nameKey = normalizeNameKey(entry.name);
const emails = [entry.email];
const ids = [entry.id];

assert.deepEqual(tokenizeSearchTerm("ojiry"), ["ojiry"]);
assert.deepEqual(tokenizeSearchTerm("ojiry yoshi"), ["ojiry", "yoshi"]);
assert.deepEqual(tokenizeSearchTerm("  ojiry   yoshi  "), ["ojiry", "yoshi"]);

assert.equal(matchesTokens(tokenizeSearchTerm("ojiry"), nameKey, emails, ids), true);
assert.equal(matchesTokens(tokenizeSearchTerm("ojiry yoshi"), nameKey, emails, ids), true);
assert.equal(matchesTokens(tokenizeSearchTerm("ojiry yishi"), nameKey, emails, ids), true);
assert.equal(matchesTokens(tokenizeSearchTerm("ojiry nope"), nameKey, emails, ids), false);
assert.equal(matchesTokens(tokenizeSearchTerm("ojir"), nameKey, emails, ids), true);

console.log("search-filter.test.js: OK");
