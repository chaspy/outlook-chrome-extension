const { test, expect } = require("../fixtures");

const MOCK_URL = "http://localhost:3737/calendar-combined.html";

test.describe("Conflict Detection", () => {
  test("conflict button appears on the page", async ({ page }) => {
    await page.goto(MOCK_URL, { waitUntil: "domcontentloaded" });
    const button = page.locator("#oce-conflict-button");
    await expect(button).toBeVisible({ timeout: 10_000 });
    await expect(button).toHaveText("重複検出");
  });

  test("clicking button highlights overlapping events", async ({ page }) => {
    await page.goto(MOCK_URL, { waitUntil: "domcontentloaded" });
    const button = page.locator("#oce-conflict-button");
    await expect(button).toBeVisible({ timeout: 10_000 });

    await button.click();

    // Events 1 and 2 overlap — they should get the conflict class
    const evt1 = page.locator("[data-calitemid='evt-1']");
    const evt2 = page.locator("[data-calitemid='evt-2']");
    await expect(evt1).toHaveClass(/oce-conflict/, { timeout: 5_000 });
    await expect(evt2).toHaveClass(/oce-conflict/);

    // Event 3 does not overlap — should NOT have conflict class
    const evt3 = page.locator("[data-calitemid='evt-3']");
    await expect(evt3).not.toHaveClass(/oce-conflict/);
  });

  test("cancelled and free events are ignored", async ({ page }) => {
    await page.goto(MOCK_URL, { waitUntil: "domcontentloaded" });
    const button = page.locator("#oce-conflict-button");
    await expect(button).toBeVisible({ timeout: 10_000 });

    await button.click();

    // Cancelled event should not be highlighted
    const evt4 = page.locator("[data-calitemid='evt-4']");
    await expect(evt4).not.toHaveClass(/oce-conflict/);

    // Free-status event should not be highlighted
    const evt5 = page.locator("[data-calitemid='evt-5']");
    await expect(evt5).not.toHaveClass(/oce-conflict/);
  });

  test("clicking again clears highlights", async ({ page }) => {
    await page.goto(MOCK_URL, { waitUntil: "domcontentloaded" });
    const button = page.locator("#oce-conflict-button");
    await expect(button).toBeVisible({ timeout: 10_000 });

    // First click: activate
    await button.click();
    const evt1 = page.locator("[data-calitemid='evt-1']");
    await expect(evt1).toHaveClass(/oce-conflict/, { timeout: 5_000 });

    // Second click: deactivate
    await button.click();
    await expect(evt1).not.toHaveClass(/oce-conflict/, { timeout: 5_000 });

    // Button text should revert
    await expect(button).toHaveText("重複検出");
  });

  test("toast shows conflict count", async ({ page }) => {
    await page.goto(MOCK_URL, { waitUntil: "domcontentloaded" });
    const button = page.locator("#oce-conflict-button");
    await expect(button).toBeVisible({ timeout: 10_000 });

    await button.click();

    const toast = page.locator("#oce-conflict-toast");
    await expect(toast).toBeVisible({ timeout: 5_000 });
    await expect(toast).toContainText("重複候補");
  });
});
