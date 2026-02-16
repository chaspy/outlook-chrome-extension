#!/usr/bin/env node

/* eslint-env node */

const fs = require("node:fs");

const VERSION_FILE = "VERSION";
const VERSION_PATTERN = /^v?(\d+)\.(\d+)\.(\d+)$/;

const readJson = (path) => JSON.parse(fs.readFileSync(path, "utf8"));
const writeJson = (path, json) => fs.writeFileSync(path, `${JSON.stringify(json, null, 2)}\n`);

const version = fs.readFileSync(VERSION_FILE, "utf8").trim();
if (!version) {
  console.error(`VERSION_SYNC_FAILED: ${VERSION_FILE} is empty.`);
  process.exit(1);
}
if (!VERSION_PATTERN.test(version)) {
  console.error(
    `VERSION_SYNC_FAILED: ${VERSION_FILE} value "${version}" is invalid. Use "X.Y.Z".`
  );
  process.exit(1);
}

const manifest = readJson("manifest.json");
const pkg = readJson("package.json");
const lock = readJson("package-lock.json");

manifest.version = version;
pkg.version = version;
lock.version = version;
lock.packages ??= {};
lock.packages[""] ??= {};
lock.packages[""].version = version;

writeJson("manifest.json", manifest);
writeJson("package.json", pkg);
writeJson("package-lock.json", lock);

console.log(`Synced manifest/package/package-lock versions to ${version}.`);
