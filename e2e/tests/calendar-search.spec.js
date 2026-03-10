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

test.describe("Calendar Search", () => {
  test("search box is visible", async ({ page, extensionId }) => {
    await seedStorage(page, extensionId, CONTACTS);
    await page.goto(MOCK_URL, { waitUntil: "domcontentloaded" });

    const searchBox = page.locator("#oce-calendar-search");
    await expect(searchBox).toBeVisible({ timeout: 10_000 });

    const searchInput = page.locator("#oce-calendar-search-input");
    await expect(searchInput).toBeVisible();
  });

  test("typing filters calendar rows", async ({ page, extensionId }) => {
    await seedStorage(page, extensionId, CONTACTS);
    await page.goto(MOCK_URL, { waitUntil: "domcontentloaded" });

    const searchInput = page.locator("#oce-calendar-search-input");
    await expect(searchInput).toBeVisible({ timeout: 10_000 });

    // Search for "田中"
    await searchInput.fill("田中");

    // Wait for search hit/miss classes to be applied
    await expect(page.locator(".oce-calendar-search-hit").first()).toBeVisible({
      timeout: 5_000
    });

    // 田中太郎 row should be a hit
    const hitRows = page.locator("button[role='option'].oce-calendar-search-hit");
    const missRows = page.locator("button[role='option'].oce-calendar-search-miss");

    const hitCount = await hitRows.count();
    const missCount = await missRows.count();

    expect(hitCount).toBeGreaterThanOrEqual(1);
    expect(missCount).toBeGreaterThanOrEqual(1);
  });

  test("searching by email filters correctly", async ({ page, extensionId }) => {
    await seedStorage(page, extensionId, CONTACTS);
    await page.goto(MOCK_URL, { waitUntil: "domcontentloaded" });

    const searchInput = page.locator("#oce-calendar-search-input");
    await expect(searchInput).toBeVisible({ timeout: 10_000 });

    // Search by email
    await searchInput.fill("tanaka@");

    // Wait for filtering
    await expect(page.locator(".oce-calendar-search-hit").first()).toBeVisible({
      timeout: 5_000
    });

    const hitRows = page.locator("button[role='option'].oce-calendar-search-hit");
    const hitCount = await hitRows.count();
    expect(hitCount).toBeGreaterThanOrEqual(1);
  });

  test("Escape clears the search", async ({ page, extensionId }) => {
    await seedStorage(page, extensionId, CONTACTS);
    await page.goto(MOCK_URL, { waitUntil: "domcontentloaded" });

    const searchInput = page.locator("#oce-calendar-search-input");
    await expect(searchInput).toBeVisible({ timeout: 10_000 });

    // Type and then press Escape
    await searchInput.fill("田中");
    await expect(page.locator(".oce-calendar-search-hit").first()).toBeVisible({
      timeout: 5_000
    });

    await searchInput.press("Escape");

    // Input should be cleared
    await expect(searchInput).toHaveValue("");

    // Search classes should be removed
    await expect(page.locator(".oce-calendar-search-hit")).toHaveCount(0, {
      timeout: 5_000
    });
    await expect(page.locator(".oce-calendar-search-miss")).toHaveCount(0);
  });
});
