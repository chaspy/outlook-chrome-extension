#!/usr/bin/env node

const fs = require("node:fs");
const { spawnSync } = require("node:child_process");

const MANIFEST_PATH = "manifest.json";
const GECKO_ID = "outlook-chrome-extension@chaspy.local";
const LINT_ARGS = [
  "lint",
  "--source-dir",
  ".",
  "--ignore-files",
  "node_modules/**",
  "dist/**",
  "coverage/**",
  "tests/**",
  ".github/**"
];

const originalText = fs.readFileSync(MANIFEST_PATH, "utf8");
const manifest = JSON.parse(originalText);

let updated = false;
manifest.browser_specific_settings ??= {};
manifest.browser_specific_settings.gecko ??= {};
if (!manifest.browser_specific_settings.gecko.id) {
  manifest.browser_specific_settings.gecko.id = GECKO_ID;
  updated = true;
}

if (updated) {
  fs.writeFileSync(MANIFEST_PATH, `${JSON.stringify(manifest, null, 2)}\n`);
}

try {
  const result = spawnSync("web-ext", LINT_ARGS, { stdio: "inherit" });
  if (result.error) throw result.error;
  process.exit(result.status ?? 1);
} finally {
  if (updated) {
    fs.writeFileSync(MANIFEST_PATH, originalText);
  }
}
