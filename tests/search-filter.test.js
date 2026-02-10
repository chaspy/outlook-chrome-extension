const assert = require("node:assert/strict");
const test = require("node:test");
const {
  tokenizeSearchTerm,
  matchesTokens,
  normalizeNameKey
} = require("../search-utils");

const entry = {
  name: "吉岡 良治",
  email: "ryoji_yoshioka@r.recruit.co.jp",
  id: "ojiry"
};

const nameKey = normalizeNameKey(entry.name);
const emails = [entry.email];
const ids = [entry.id];

const matches = (input) =>
  matchesTokens(tokenizeSearchTerm(input), {
    nameKey,
    emails,
    ids
  });

test("tokenizeSearchTerm splits tokens", () => {
  assert.deepEqual(tokenizeSearchTerm("ojiry"), ["ojiry"]);
  assert.deepEqual(tokenizeSearchTerm("ojiry yoshi"), ["ojiry", "yoshi"]);
  assert.deepEqual(tokenizeSearchTerm("  ojiry   yoshi  "), ["ojiry", "yoshi"]);
  assert.deepEqual(tokenizeSearchTerm("ojiry\u3000yoshi"), ["ojiry", "yoshi"]);
});

test("matchesTokens matches id and email", () => {
  assert.equal(matches("ojiry"), true);
  assert.equal(matches("ojiry yoshi"), true);
  assert.equal(matches("ojiry yishi"), false);
  assert.equal(matches("ojiry nope"), false);
  assert.equal(matches("ojir"), true);
});
