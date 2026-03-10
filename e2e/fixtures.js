const path = require("node:path");
const fs = require("node:fs");
const os = require("node:os");
const { test: base, chromium } = require("@playwright/test");

const EXT_SRC = path.resolve(__dirname, "..");

/**
 * Files that comprise the extension (no sub-directories needed).
 */
const EXT_FILES = [
  "manifest.json",
  "content.js",
  "content.css",
  "search-utils.js",
  "popup.html",
  "popup.js",
  "popup.css"
];

/**
 * Copy the extension to a temp directory and patch manifest.json
 * so that content scripts also run on http://localhost:3737/*.
 * Also adds a minimal service worker so Playwright can detect the extension ID.
 */
function prepareExtension() {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "oce-e2e-"));

  for (const file of EXT_FILES) {
    const src = path.join(EXT_SRC, file);
    if (fs.existsSync(src)) {
      fs.copyFileSync(src, path.join(tmpDir, file));
    }
  }

  // Create a minimal service worker for extension ID detection
  fs.writeFileSync(path.join(tmpDir, "background.js"), "// e2e stub\n");

  const manifestPath = path.join(tmpDir, "manifest.json");
  const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf-8"));

  const localhostPattern = "http://localhost:3737/*";

  // Patch content_scripts matches
  if (manifest.content_scripts) {
    for (const cs of manifest.content_scripts) {
      if (!cs.matches.includes(localhostPattern)) {
        cs.matches.push(localhostPattern);
      }
    }
  }

  // Patch host_permissions
  if (!manifest.host_permissions) {
    manifest.host_permissions = [];
  }
  if (!manifest.host_permissions.includes(localhostPattern)) {
    manifest.host_permissions.push(localhostPattern);
  }

  // Add background service worker for extension ID detection
  if (!manifest.background) {
    manifest.background = { service_worker: "background.js" };
  }

  fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2));

  return tmpDir;
}

/**
 * Custom Playwright fixtures that launch Chromium with the extension loaded.
 */
const test = base.extend({
  // eslint-disable-next-line no-empty-pattern
  context: async ({}, use) => {
    const extDir = prepareExtension();
    const context = await chromium.launchPersistentContext("", {
      headless: false,
      args: [
        `--disable-extensions-except=${extDir}`,
        `--load-extension=${extDir}`,
        "--no-first-run",
        "--disable-default-apps",
        "--disable-popup-blocking"
      ]
    });

    await use(context);

    await context.close();
    fs.rmSync(extDir, { recursive: true, force: true });
  },

  extensionId: async ({ context }, use) => {
    let serviceWorker;
    if (context.serviceWorkers().length > 0) {
      serviceWorker = context.serviceWorkers()[0];
    } else {
      serviceWorker = await context.waitForEvent("serviceworker");
    }
    const extId = serviceWorker.url().split("/")[2];
    await use(extId);
  },

  page: async ({ context }, use) => {
    const page = await context.newPage();
    await use(page);
  }
});

const { expect } = require("@playwright/test");
module.exports = { test, expect };
