const { test, expect } = require("../fixtures");
const { seedStorage } = require("../helpers/seed-storage");

const MOCK_URL = "http://localhost:3737/calendar-combined.html";

const CONTACTS = [
  { name: "田中太郎", email: "tanaka@example.com", id: "tanaka01" },
  { name: "鈴木花子", email: "suzuki@example.com", id: "suzuki01" },
  { name: "佐藤次郎", email: "sato@example.com", id: "sato01" },
  { name: "山田三郎", email: "yamada@example.com", id: "yamada01" },
  { name: "高橋美咲", email: "takahashi@example.com", id: "takahashi01" },
  { name: "渡辺健一", email: "watanabe@example.com", id: "watanabe01" }
];

test.describe("Copy Buttons", () => {
  test("selected calendar names show copy buttons", async ({ page, extensionId }) => {
    await seedStorage(page, extensionId, CONTACTS);
    await page.goto(MOCK_URL, { waitUntil: "domcontentloaded" });

    // Wait for the selected summary to appear
    const summary = page.locator("#oce-calendar-selected-summary");
    await expect(summary).toBeVisible({ timeout: 10_000 });

    // Selected calendars (田中太郎, 鈴木花子, 高橋美咲) should have copy buttons
    const copyButtons = summary.locator(".oce-copy-btn");
    // Each selected entry gets name copy + email copy = 2 buttons per entry
    // 3 selected calendars × 2 = 6 copy buttons
    await expect(copyButtons.first()).toBeVisible({ timeout: 5_000 });
    const count = await copyButtons.count();
    expect(count).toBeGreaterThanOrEqual(3);
  });

  test("clicking copy button shows toast", async ({ page, extensionId }) => {
    await seedStorage(page, extensionId, CONTACTS);
    await page.goto(MOCK_URL, { waitUntil: "domcontentloaded" });

    const summary = page.locator("#oce-calendar-selected-summary");
    await expect(summary).toBeVisible({ timeout: 10_000 });

    // Click the first copy button
    const firstCopyBtn = summary.locator(".oce-copy-btn").first();
    await expect(firstCopyBtn).toBeVisible({ timeout: 5_000 });
    await firstCopyBtn.click();

    // Toast should appear with "コピーしました"
    const toast = page.locator("#oce-conflict-toast");
    await expect(toast).toBeVisible({ timeout: 5_000 });
    await expect(toast).toContainText("コピーしました");
  });

  test("search hint results also have copy buttons", async ({ page, extensionId }) => {
    await seedStorage(page, extensionId, CONTACTS);
    await page.goto(MOCK_URL, { waitUntil: "domcontentloaded" });

    // Wait for search box
    const searchInput = page.locator("#oce-calendar-search-input");
    await expect(searchInput).toBeVisible({ timeout: 10_000 });

    // Type a search term
    await searchInput.fill("田中");

    // Wait for search hints to appear
    const hints = page.locator("#oce-calendar-search-hints");
    await expect(hints).toBeVisible({ timeout: 5_000 });

    // Hints should have copy buttons
    const hintCopyButtons = hints.locator(".oce-copy-btn");
    await expect(hintCopyButtons.first()).toBeVisible({ timeout: 5_000 });

    // Click a copy button in hints
    await hintCopyButtons.first().click();

    const toast = page.locator("#oce-conflict-toast");
    await expect(toast).toBeVisible({ timeout: 5_000 });
    await expect(toast).toContainText("コピーしました");
  });
});
