/**
 * Seed chrome.storage.local with contact data via the extension's popup.html.
 *
 * @param {import('@playwright/test').Page} page - Playwright page (persistent context)
 * @param {string} extensionId - The loaded extension's ID
 * @param {Array<{name:string, email:string, id?:string}>} contacts
 */
async function seedStorage(page, extensionId, contacts) {
  const popupUrl = `chrome-extension://${extensionId}/popup.html`;
  await page.goto(popupUrl, { waitUntil: "domcontentloaded" });

  await page.evaluate((data) => {
    return new Promise((resolve, reject) => {
      chrome.storage.local.set({ oceContacts: data }, () => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          resolve();
        }
      });
    });
  }, contacts);
}

module.exports = { seedStorage };
