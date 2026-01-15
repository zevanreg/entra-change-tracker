/**
 * Entra Portal Scraper
 * Handles all browser automation and data extraction from Entra portal
 */

const { chromium } = require("playwright");
const { waitForSplashScreen, clickTab, setDateRangeFilter, scrapeDetailsList } = require("./browser-helpers");

const ENTRA_URL = "https://entra.microsoft.com/#blade/Microsoft_AAD_IAM/ChangeManagementHubList.ReactView";

/**
 * Initialize browser and navigate to Entra portal
 * @returns {Promise<{context: import('playwright').BrowserContext, page: import('playwright').Page, frame: import('playwright').Frame}>}
 */
async function initializeBrowser() {
  const context = await chromium.launchPersistentContext('./edge-profile', {
    channel: 'msedge',
    headless: false,
    args: ['--disable-blink-features=AutomationControlled'],
    viewport: { width: 1920, height: 1080 },
    deviceScaleFactor: 1,
  });
  
  const page = context.pages()[0];

  // Navigate to Entra portal
  await page.goto(ENTRA_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });

  // Get the main iframe
  const iframeLocator = page.locator('iframe[name="ChangeManagementHubList.ReactView"]');
  await iframeLocator.waitFor({ state: 'attached', timeout: 30000 });

  const iframeHandle = await iframeLocator.elementHandle();
  const frame = await iframeHandle.contentFrame();
  
  if (!frame) {
    throw new Error('ReactView frame attached, but content not available yet.');
  }

  // Wait for initial splash screen to disappear
  await waitForSplashScreen(frame, 60000);

  // Wait for progress dots to disappear
  try {
    const progressDots = frame.locator('div.fxs-progress-dots');
    await progressDots.waitFor({ state: 'hidden', timeout: 15000 });
  } catch (err) {
    // Progress dots might not exist or already hidden
  }

  return { context, page, frame };
}

/**
 * Scrape data from a specific tab
 * @param {import('playwright').Page} page - The page
 * @param {import('playwright').Frame} frame - The iframe
 * @param {string|RegExp} tabName - The name of the tab to scrape
 * @param {string|null} dateFilter - Optional date filter to apply
 * @returns {Promise<Array<Object>|null>} Scraped data or null if tab not found
 */
async function scrapeTab(page, frame, tabName, dateFilter = null) {
  // Click the tab
  const tabClicked = await clickTab(frame, tabName);

  if (!tabClicked) {
    console.error(`❌ Could not locate the ${tabName} tab/menu.`);
    return null;
  }

  // Set date filter if specified
  if (dateFilter) {
    const filterSet = await setDateRangeFilter(frame, dateFilter);
    if (!filterSet) {
      console.warn('⚠️ Could not set date range filter, continuing anyway...');
    }
  }

  // Scrape the data
  const data = await scrapeDetailsList(page, frame);
  console.log(`✅ Extracted ${data.length} items from ${tabName}`);

  return data;
}

/**
 * Scrape Roadmap data from Entra portal
 * @param {string|null} dateFilter - Optional date filter
 * @returns {Promise<{roadmap: Array<Object>|null, changeAnnouncements: Array<Object>|null}>}
 */
async function scrapeEntraPortal(dateFilter = null) {
  let context, page, frame;
  
  try {
    // Initialize browser and navigate
    ({ context, page, frame } = await initializeBrowser());

    // Scrape Roadmap
    const roadmap = await scrapeTab(page, frame, /^Roadmap$/i, dateFilter);

    // Scrape Change Announcements
    const changeAnnouncements = await scrapeTab(page, frame, /^Change announcements$/i, dateFilter);

    return { roadmap, changeAnnouncements };
  } finally {
    if (context) {
      await context.close();
      console.log('✅ Browser closed.');
    }
  }
}

module.exports = {
  initializeBrowser,
  scrapeTab,
  scrapeEntraPortal,
};
