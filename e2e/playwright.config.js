const { defineConfig } = require("@playwright/test");
const path = require("node:path");

module.exports = defineConfig({
  testDir: path.join(__dirname, "tests"),
  timeout: 30_000,
  retries: 1,
  workers: 1,
  reporter: [["html", { outputFolder: path.join(__dirname, "..", "playwright-report") }]],
  outputDir: path.join(__dirname, "..", "test-results"),
  use: {
    headless: false,
    viewport: { width: 1280, height: 800 }
  },
  webServer: {
    command: `node ${path.join(__dirname, "helpers", "serve.js")}`,
    port: 3737,
    reuseExistingServer: !process.env.CI
  }
});
