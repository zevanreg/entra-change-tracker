/**
 * Entra Portal Scraper
 * Handles all browser automation and data extraction from Entra portal
 */

const { chromium } = require("playwright");
const axios = require("axios");
const cheerio = require("cheerio");
const { waitForSplashScreen, clickTab, setDateRangeFilter, scrapeDetailsList } = require("./browser-helpers");
const { getConfiguration } = require("./auth");

const ENTRA_URL = "https://entra.microsoft.com/#blade/Microsoft_AAD_IAM/ChangeManagementHubList.ReactView";

/**
 * Extract release type from the beginning of the title
 * @param {string} title - The full title string
 * @returns {{releaseType: string, cleanedTitle: string}}
 */
function extractReleaseTypeFromTitle(title) {
  const { config } = getConfiguration();
  const releaseTypeMapping = config?.releaseTypeMapping || {};
  
  // Check if title starts with any known release type (case-insensitive)
  const titleLower = title.toLowerCase();
  for (const [pageValue, mappedValue] of Object.entries(releaseTypeMapping)) {
    if (titleLower.startsWith(pageValue.toLowerCase())) {
      // Extract the part after the release type
      let remainder = title.slice(pageValue.length).trim();
      // Remove common separators at the start
      const separators = ['-', ':', '–', '—'];
      for (const sep of separators) {
        if (remainder.startsWith(sep)) {
          remainder = remainder.slice(1).trim();
          break;
        }
      }
      return { releaseType: mappedValue, cleanedTitle: remainder };
    }
  }
  
  // No release type found - check if there's an unmapped one
  if ([' - ', ': ', ' – '].some(sep => title.slice(0, 50).includes(sep))) {
    for (const sep of [' - ', ': ', ' – ']) {
      if (title.includes(sep)) {
        const potentialType = title.split(sep)[0].trim();
        if (potentialType && /^[A-Z]/.test(potentialType)) {
          console.warn(`⚠️ Unmapped release type found: '${potentialType}' in title: ${title.slice(0, 60)}...`);
        }
        break;
      }
    }
  }
  
  return { releaseType: '', cleanedTitle: title };
}

/**
 * Extract a single item from the What's New page
 * @param {import('cheerio').Cheerio} h3Element - The h3 element
 * @param {string} monthText - The month/year text
 * @param {import('cheerio').CheerioAPI} $ - Cheerio instance
 * @returns {Object|null}
 */
function extractWhatsNewItem(h3Element, monthText, $) {
  try {
    const item = {
      releaseType: '',
      title: '',
      type: '',
      serviceCategory: '',
      productCapability: '',
      detail: '',
      link: '',
      date: ''
    };
    
    // Extract title and link
    const titleLink = h3Element.find('a');
    let fullTitle;
    if (titleLink.length > 0) {
      fullTitle = titleLink.text().trim();
      item.link = titleLink.attr('href') || '';
      if (item.link && !item.link.startsWith('http')) {
        item.link = `https://learn.microsoft.com${item.link}`;
      }
    } else {
      fullTitle = h3Element.text().trim();
    }
    
    // Extract release type from title
    const { releaseType, cleanedTitle } = extractReleaseTypeFromTitle(fullTitle);
    
    if (!releaseType && [' - ', ': ', ' – '].some(sep => fullTitle.includes(sep))) {
      for (const sep of [' - ', ': ', ' – ']) {
        if (fullTitle.includes(sep)) {
          const potentialType = fullTitle.split(sep)[0].trim();
          if (potentialType && /^[A-Z]/.test(potentialType)) {
            console.warn(`⚠️ Unmapped release type found: '${potentialType}' in title: ${fullTitle.slice(0, 60)}...`);
          }
          break;
        }
      }
    }
    
    item.releaseType = releaseType;
    item.title = cleanedTitle;
    item.date = monthText;
    
    // Extract detail from following paragraphs
    const detailParts = [];
    let current = h3Element.next();
    
    while (current.length > 0 && !['h2', 'h3'].includes(current.prop('tagName')?.toLowerCase())) {
      if (current.prop('tagName')?.toLowerCase() === 'p') {
        const text = current.text().trim();
        if (text) {
          // Look for metadata in strong tags
          current.find('strong').each((_, strong) => {
            const label = $(strong).text().trim().replace(/:$/, '');
            const nextText = $(strong)[0]?.nextSibling;
            if (nextText && nextText.type === 'text') {
              const value = nextText.data.trim().replace(/^:/, '');
              
              if (label.includes('Type') && !item.type) {
                item.type = value;
              } else if (label.toLowerCase().includes('service category')) {
                item.serviceCategory = value;
              } else if (label.toLowerCase().includes('product capability')) {
                item.productCapability = value;
              }
            }
          });
          
          detailParts.push(text);
        }
      } else if (current.prop('tagName')?.toLowerCase() === 'ul') {
        current.find('li').each((_, li) => {
          detailParts.push(`• ${$(li).text().trim()}`);
        });
      }
      
      current = current.next();
    }
    
    item.detail = detailParts.join(' ');
    
    return item.title ? item : null;
  } catch (error) {
    console.warn(`⚠️ Error extracting item: ${error.message}`);
    return null;
  }
}

/**
 * Scrape data from Microsoft Learn What's New page
 * @returns {Promise<Array<Object>|null>}
 */
async function scrapeWhatsNewPage() {
  const url = "https://learn.microsoft.com/en-us/entra/fundamentals/whats-new";
  
  try {
    console.log(`🌐 Fetching ${url}...`);
    const response = await axios.get(url, { timeout: 30000 });
    
    const $ = cheerio.load(response.data);
    const items = [];
    
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                    'July', 'August', 'September', 'October', 'November', 'December'];
    
    // Find all h2 headers that represent month sections
    $('h2[id]').each((_, monthSection) => {
      const monthText = $(monthSection).text().trim();
      
      // Skip if not a date header
      if (!months.some(month => monthText.includes(month))) {
        return;
      }
      
      // Find all h3 items under this month
      let current = $(monthSection).next();
      while (current.length > 0 && current.prop('tagName')?.toLowerCase() !== 'h2') {
        if (current.prop('tagName')?.toLowerCase() === 'h3') {
          const item = extractWhatsNewItem(current, monthText, $);
          if (item) {
            items.push(item);
          }
        }
        current = current.next();
      }
    });
    
    console.log(`✅ Extracted ${items.length} items from What's New page`);
    return items;
  } catch (error) {
    console.error(`❌ Error scraping What's New page: ${error.message}`);
    return null;
  }
}

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

/**
 * Scrape data from all sources: Entra portal and Microsoft Learn What's New
 * @param {string|null} dateFilter - Optional date filter for portal scraping
 * @returns {Promise<{roadmap: Array<Object>|null, changeAnnouncements: Array<Object>|null, whatsNew: Array<Object>|null}>}
 */
async function scrapeAllSources(dateFilter = null) {
  // Scrape Entra portal data
  const portalData = await scrapeEntraPortal(dateFilter);
  
  // Scrape What's New page
  console.log("\n📚 Scraping Microsoft Learn What's New page...");
  const whatsNew = await scrapeWhatsNewPage();
  
  return {
    roadmap: portalData.roadmap,
    changeAnnouncements: portalData.changeAnnouncements,
    whatsNew
  };
}

module.exports = {
  initializeBrowser,
  scrapeTab,
  scrapeEntraPortal,
  scrapeAllSources,
};
