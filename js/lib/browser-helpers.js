/**
 * Browser automation helper functions for Entra portal scraping
 */

const { getConfig } = require('./config');

// ==================== LAZY CONFIG GETTERS ====================

let _browserConfig = null;

function getBrowserConfig() {
  if (!_browserConfig) {
    const config = getConfig();
    if (!config) {
      throw new Error('Configuration not loaded. Call loadConfiguration() first.');
    }
    _browserConfig = config.browserScraping;
  }
  return _browserConfig;
}

// ==================== CONSTANTS ====================

function getTimeouts() {
  const browserConfig = getBrowserConfig();
  return {
    SPLASH_SCREEN: browserConfig.timeouts.splashScreen,
    PROGRESS_DOTS: browserConfig.timeouts.progressDots,
    GENERAL_WAIT: browserConfig.timeouts.generalWait,
    CLICK: browserConfig.timeouts.click,
    DETACH: browserConfig.timeouts.detach,
    CLOSE_PANE: browserConfig.timeouts.closePane,
    BUTTON_CLOSE: browserConfig.timeouts.buttonClose,
    SHORT_DELAY: browserConfig.timeouts.shortDelay,
    MENU_DELAY: browserConfig.timeouts.menuDelay,
    CHECKBOX_DELAY: browserConfig.timeouts.checkboxDelay
  };
}

function getSelectors() {
  const browserConfig = getBrowserConfig();
  return {
    SPLASH_SCREEN: browserConfig.selectors.splashScreen,
    DETAILS_IFRAME: browserConfig.selectors.detailsIframe,
    DETAILS_ROW: browserConfig.selectors.detailsRow,
    DETAILS_ROW_CHECK: browserConfig.selectors.detailsRowCheck,
    DETAILS_ROW_FIELDS: browserConfig.selectors.detailsRowFields,
    DETAILS_ROW_CELL: browserConfig.selectors.detailsRowCell,
    PROGRESS_DOTS: browserConfig.selectors.progressDots,
    CLOSE_BUTTON: browserConfig.selectors.closeButton,
    SCROLLABLE_CONTAINER: browserConfig.selectors.scrollableContainer,
    FILTER_BUTTON_CONTAINER: browserConfig.selectors.filterButtonContainer,
    APPLY_BUTTON: browserConfig.selectors.applyButton,
    RADIO_LABEL: browserConfig.selectors.radioLabel
  };
}

function getScraperConfig() {
  const browserConfig = getBrowserConfig();
  return {
    SCROLL_STEP_PX: browserConfig.scraperConfig.scrollStepPx,
    PASS_DELAY_MS: browserConfig.scraperConfig.passDelayMs,
    MAX_IDLE_PASSES: browserConfig.scraperConfig.maxIdlePasses,
    TOTAL_TIMEOUT_MS: browserConfig.scraperConfig.totalTimeoutMs,
    MAX_RETRY_ATTEMPTS: browserConfig.scraperConfig.maxRetryAttempts,
    CLICK_OUTSIDE_COORDS: browserConfig.scraperConfig.clickOutsideCoords
  };
}

function getTextPatterns() {
  const browserConfig = getBrowserConfig();
  return {
    OVERVIEW: browserConfig.textPatterns.overview,
    NEXT_STEPS: browserConfig.textPatterns.nextSteps,
    WHAT_IS_CHANGING: browserConfig.textPatterns.whatIsChanging,
    ROADMAP_DESCRIPTION: browserConfig.textPatterns.roadmapDescription
  };
}

// ==================== HELPER FUNCTIONS ====================

/**
 * Waits for the Entra splash screen to disappear
 * @param {import('playwright').Frame} frame - The frame to check
 * @param {number} timeout - Maximum wait time in ms
 */
async function waitForSplashScreen(frame, timeout) {
  const TIMEOUTS = getTimeouts();
  const SELECTORS = getSelectors();
  const actualTimeout = timeout || TIMEOUTS.SPLASH_SCREEN;
  
  try {
    const splashScreen = frame.locator(SELECTORS.SPLASH_SCREEN);
    await splashScreen.waitFor({ state: 'hidden', timeout: actualTimeout });
  } catch (err) {
    // Splash screen might not exist or already hidden
    console.log('Splash screen not found or already hidden');
  }
}

/**
 * Clicks on a tab within a frame using role-based locators
 * @param {import('playwright').Frame} frame - The frame containing the tab
 * @param {string|RegExp} tabName - The name of the tab to click (string or regex)
 * @returns {Promise<boolean>} True if the tab was found and clicked, false otherwise
 */
async function clickTab(frame, tabName) {
  const TIMEOUTS = getTimeouts();
  
  try {
    // Wait for splash screen to disappear first
    await waitForSplashScreen(frame);
    
    const roleTab = frame.getByRole('tab', { name: tabName });
    const count = await roleTab.count();
    if (count > 0) {
      // Wait for tab to be visible and stable
      await roleTab.first().waitFor({ state: 'visible', timeout: TIMEOUTS.GENERAL_WAIT });
      
      try {
        // Try normal click first
        await roleTab.first().click({ timeout: TIMEOUTS.CLICK });
      } catch (clickErr) {
        // If intercepted, force click
        console.log(`Normal click failed, forcing click on tab "${tabName}"`);
        await roleTab.first().click({ force: true, timeout: TIMEOUTS.CLICK });
      }
      
      // Wait a moment for tab content to load
      await frame.waitForTimeout(TIMEOUTS.SHORT_DELAY);
      return true;
    }
  } catch (err) {
    console.error(`Error clicking tab "${tabName}":`, err.message);
  }
  return false;
}

/**
 * Sets the date range filter to a specific option
 * @param {import('playwright').Frame} frame - The frame containing the filter
 * @param {string} filterOption - The text of the filter option to select (e.g., "Last 1 month")
 * @returns {Promise<boolean>} True if the filter was set successfully
 */
async function setDateRangeFilter(frame, filterOption) {
  const TIMEOUTS = getTimeouts();
  const SELECTORS = getSelectors();
  
  try {
    // Click the filter button - find button inside div with data-selection-index='1'
    const filterButton = frame.locator(`${SELECTORS.FILTER_BUTTON_CONTAINER} button`).first();
    await filterButton.waitFor({ state: 'visible', timeout: TIMEOUTS.GENERAL_WAIT });
    await filterButton.click({ timeout: TIMEOUTS.CLICK });
    console.log('Filter button clicked');
    
    // Wait for the filter menu to appear
    await frame.waitForTimeout(TIMEOUTS.MENU_DELAY);
    
    // Find and click the radio input with the matching label
    const radioLabel = frame.locator(`span${SELECTORS.RADIO_LABEL}:has-text("${filterOption}")`);
    await radioLabel.waitFor({ state: 'visible', timeout: TIMEOUTS.CLICK });
    
    // Click the label to select the radio button
    await radioLabel.click({ timeout: TIMEOUTS.CLICK });
    console.log(`Selected filter option: ${filterOption}`);
    
    // Wait a moment for the selection to register
    await frame.waitForTimeout(TIMEOUTS.MENU_DELAY);
    
    // Click the Apply button
    const applyButton = frame.locator(SELECTORS.APPLY_BUTTON);
    if (await applyButton.count() > 0 && !await applyButton.isDisabled()) {
      await applyButton.click({ timeout: TIMEOUTS.CLICK });
      console.log('Apply button clicked');
      await frame.waitForTimeout(TIMEOUTS.SHORT_DELAY);
    }
    
    return true;
  } catch (err) {
    console.error(`Error setting date range filter to "${filterOption}":`, err.message);
    return false;
  }
}

// ==================== DETAILS EXTRACTION HELPERS ====================

/**
 * Extracts overview text from the details pane
 * @param {import('playwright').Frame} detailsFrame - The details iframe
 * @param {number} rowIndex - Row index for logging
 * @returns {Promise<string>} The extracted overview text
 */
async function extractOverview(detailsFrame, rowIndex) {
  const TEXT_PATTERNS = getTextPatterns();
  
  try {
    if (detailsFrame.isDetached()) return '';
    
    const overviewH3 = detailsFrame.locator('h3').filter({ hasText: TEXT_PATTERNS.OVERVIEW });
    
    if (await overviewH3.count() > 0) {
      const parent = overviewH3.locator('..');
      const spans = parent.locator('span');
      
      if (await spans.count() > 0) {
        return await spans.first().innerText() || '';
      }
    }
  } catch (err) {
    console.warn(`Could not extract overview for row ${rowIndex}:`, err.message);
  }
  return '';
}

/**
 * Extracts URL from the "Next steps" section
 * @param {import('playwright').Frame} detailsFrame - The details iframe
 * @param {number} rowIndex - Row index for logging
 * @returns {Promise<string>} The extracted URL
 */
async function extractUrl(detailsFrame, rowIndex) {
  const TEXT_PATTERNS = getTextPatterns();
  
  try {
    if (detailsFrame.isDetached()) return '';
    
    const nextStepsH3 = detailsFrame.locator('h3').filter({ hasText: TEXT_PATTERNS.NEXT_STEPS });
    
    if (await nextStepsH3.count() > 0) {
      const parent = nextStepsH3.locator('..');
      const links = parent.locator('a');
      
      if (await links.count() > 0) {
        return await links.first().getAttribute('href') || '';
      }
    }
  } catch (err) {
    console.warn(`Could not extract URL for row ${rowIndex}:`, err.message);
  }
  return '';
}

/**
 * Extracts description text from either "What is changing" or roadmap section
 * @param {import('playwright').Frame} detailsFrame - The details iframe
 * @param {number} rowIndex - Row index for logging
 * @returns {Promise<string>} The extracted description text
 */
async function extractDescription(detailsFrame, rowIndex) {
  const TEXT_PATTERNS = getTextPatterns();
  
  try {
    if (detailsFrame.isDetached()) return '';
    
    // Try "What is changing" first (change announcements)
    let descH3 = detailsFrame.locator('h3').filter({ hasText: TEXT_PATTERNS.WHAT_IS_CHANGING });
    
    if (await descH3.count() > 0) {
      // Get the span that is the next sibling of the h3
      const nextSpan = descH3.locator('xpath=following-sibling::span[1]');
      
      if (await nextSpan.count() > 0) {
        return await nextSpan.innerText() || '';
      }
    } else {
      // Fall back to "Here's what you will see in this release:" (roadmap)
      descH3 = detailsFrame.locator('h3').filter({ hasText: TEXT_PATTERNS.ROADMAP_DESCRIPTION });
      
      if (await descH3.count() > 0) {
        const parent = descH3.locator('..');
        const paragraphs = parent.locator('p');
        
        if (await paragraphs.count() > 0) {
          return await paragraphs.first().innerText() || '';
        }
      }
    }
  } catch (err) {
    console.warn(`Could not extract description for row ${rowIndex}:`, err.message);
  }
  return '';
}

// ==================== IFRAME MANAGEMENT ====================

/**
 * Opens the details pane by clicking a row's checkbox
 * @param {import('playwright').Page} page - The main page
 * @param {import('playwright').Frame} frame - The frame containing the row
 * @param {number} rowIndex - The index of the row to click
 * @returns {Promise<void>}
 */
async function openDetailsPane(page, frame, rowIndex) {
  const TIMEOUTS = getTimeouts();
  const SELECTORS = getSelectors();
  
  const row = frame.locator(`${SELECTORS.DETAILS_ROW}[data-item-index='${rowIndex}']`).first();
  const checkbox = row.locator(SELECTORS.DETAILS_ROW_CHECK);
  
  // Check if the row is already selected (checked)
  const isChecked = await checkbox.getAttribute('aria-checked');
  if (isChecked === 'true') {
    // Uncheck it first by clicking
    await checkbox.click({ timeout: TIMEOUTS.CLICK });
    await page.waitForTimeout(TIMEOUTS.CHECKBOX_DELAY);
  }
  
  // Now click to select and open details pane
  await checkbox.click({ timeout: TIMEOUTS.CLICK });
  
  // Wait for at least one details iframe to appear
  await page.locator(`${SELECTORS.DETAILS_IFRAME}:visible`).first().waitFor({ 
    state: 'attached', 
    timeout: TIMEOUTS.GENERAL_WAIT 
  });
}

/**
 * Finds the correct details iframe by matching the row title
 * @param {import('playwright').Page} page - The main page
 * @param {string} rowTitle - The title to match
 * @param {number} rowIndex - Row index for logging
 * @returns {Promise<import('playwright').Frame|null>} The matched iframe or null
 */
async function findCorrectIframe(page, rowTitle, rowIndex) {
  const TIMEOUTS = getTimeouts();
  const SELECTORS = getSelectors();
  
  const allIframes = await page.locator(SELECTORS.DETAILS_IFRAME).all();
  
  for (const iframe of allIframes) {
    try {
      await iframe.waitFor({ state: 'attached', timeout: TIMEOUTS.GENERAL_WAIT });
      
      const iframeHandle = await iframe.elementHandle();
      const frame = await iframeHandle.contentFrame();
      
      if (!frame) continue;
      
      // Wait for progress dots to disappear
      try {
        const progressDots = frame.locator(SELECTORS.PROGRESS_DOTS);
        await progressDots.waitFor({ state: 'hidden', timeout: TIMEOUTS.PROGRESS_DOTS });
      } catch (err) {
        // Progress dots might not exist or already hidden
      }
      
      await frame.waitForLoadState('domcontentloaded');

      if (rowTitle) {
        // Check if this iframe contains an h3 with the row title
        const titleH3 = frame.locator('h3').filter({ hasText: rowTitle });
        if (await titleH3.count() > 0) {
          return frame;
        }
      } else {
        // If we couldn't get the title, use this iframe
        return frame;
      }
    } catch (err) {
      // Skip this iframe if we can't access it
      continue;
    }
  }
  
  console.warn(`Could not find iframe with matching title for row ${rowIndex}`);
 
  return null;
}

/**
 * Closes the details pane and waits for iframe removal
 * @param {import('playwright').Page} page - The main page
 * @param {number} rowIndex - Row index for logging
 * @returns {Promise<void>}
 */
async function closeDetailsPane(page, rowIndex) {
  const TIMEOUTS = getTimeouts();
  const SELECTORS = getSelectors();
  const SCRAPER_CONFIG = getScraperConfig();
  
  try {
    // Method 1: Look for close button in the main page (Azure blade close button)
    const closeButton = page.locator(SELECTORS.CLOSE_BUTTON).last();
    if (await closeButton.count() > 0 && await closeButton.isVisible()) {
      await closeButton.click({ timeout: TIMEOUTS.BUTTON_CLOSE });
    } else {
      // Method 2: Click outside the iframe to close it
      await page.mouse.click(SCRAPER_CONFIG.CLICK_OUTSIDE_COORDS.x, SCRAPER_CONFIG.CLICK_OUTSIDE_COORDS.y);
    }
    
    // Wait for all details iframes to be removed from DOM
    await page.waitForFunction(
      (selector) => document.querySelectorAll(selector).length === 0,
      SELECTORS.DETAILS_IFRAME,
      { timeout: TIMEOUTS.CLOSE_PANE }
    );
  } catch (err) {
    // If close button/click didn't work, try ESC key as fallback
    console.error(`Could not close details pane with button/click for row ${rowIndex}, trying ESC key...`, err.message);
    try {
      await page.keyboard.press('Escape');
      await page.waitForFunction(
        (selector) => document.querySelectorAll(selector).length === 0,
        SELECTORS.DETAILS_IFRAME,
        { timeout: TIMEOUTS.CLOSE_PANE }
      );
    } catch (escErr) {
      console.warn(`Details iframe did not detach for row ${rowIndex}, continuing anyway...`);
    }
  }
}

// ==================== MAIN EXTRACTION FUNCTION ====================

/**
 * Extracts details from a row by clicking it and scraping the details pane
 * @param {import('playwright').Page} page - The main page
 * @param {import('playwright').Frame} frame - The frame containing the row
 * @param {number} rowIndex - The index of the row to click
 * @param {string} rowTitle - The title of the row (for iframe matching)
 * @returns {Promise<{url: string, description: string, overview: string}>} The extracted details
 */
async function extractRowDetails(page, frame, rowIndex, rowTitle = '') {
  const TIMEOUTS = getTimeouts();
  const SELECTORS = getSelectors();
  const SCRAPER_CONFIG = getScraperConfig();
  
  const startTime = Date.now();
  console.log(`[${new Date().toISOString()}] 🔵 START extracting details for row ${rowIndex}`);
  
  try {
    // Open the details pane
    await openDetailsPane(page, frame, rowIndex);
    
    // Find the correct iframe
    const detailsFrame = await findCorrectIframe(page, rowTitle, rowIndex);
    
    if (!detailsFrame) {
      console.warn(`Details frame not available for row ${rowIndex}`);
      return { url: '', description: '', overview: '' };
    }
    
    // Extract all details using helper functions
    const overview = await extractOverview(detailsFrame, rowIndex);
    const url = await extractUrl(detailsFrame, rowIndex);
    const description = await extractDescription(detailsFrame, rowIndex);
    
    // Close the details pane
    await closeDetailsPane(page, rowIndex);
    
    const elapsed = Date.now() - startTime;
    console.log(`[${new Date().toISOString()}] ✅ END extracting details for row ${rowIndex} (took ${elapsed}ms)`);
    
    return { 
      url: url.trim(), 
      description: description.trim(), 
      overview: overview.trim() 
    };
  } catch (err) {
    const elapsed = Date.now() - startTime;
    console.error(`[${new Date().toISOString()}] ❌ ERROR extracting details for row ${rowIndex} (took ${elapsed}ms):`, err.message);
    
    // Attempt to close any open pane
    try {
      const detailsIframeLocator = page.locator(SELECTORS.DETAILS_IFRAME).last();
      
      // Try close button first
      const closeButton = page.locator('button[aria-label="Close"]').last();
      if (await closeButton.count() > 0) {
        await closeButton.click({ timeout: TIMEOUTS.CLOSE_PANE });
      } else {
        // Try clicking outside
        await page.mouse.click(SCRAPER_CONFIG.CLICK_OUTSIDE_COORDS.x, SCRAPER_CONFIG.CLICK_OUTSIDE_COORDS.y);
      }
      
      await detailsIframeLocator.waitFor({ state: 'detached', timeout: TIMEOUTS.DETACH });
    } catch {}
    
    return { url: '', description: '', overview: '' };
  }
}

// ==================== LIST SCRAPING ====================

/**
 * Scrapes all rows from a Fluent UI DetailsList with virtualized scrolling
 * @param {import('playwright').Page} page - The main page
 * @param {import('playwright').Frame} frame - The frame containing the DetailsList
 * @param {boolean} extractDetails - Whether to extract details (URL, description, overview) by clicking each row
 * @returns {Promise<Array<Object>>} Array of row objects with field names and details
 */
async function scrapeDetailsList(page, frame, extractDetails = true) {
  const SELECTORS = getSelectors();
  const SCRAPER_CONFIG = getScraperConfig();
  
  // Accumulator for processed rows
  const processedRows = [];
  const seenIndices = new Set();

  let idlePasses = 0;
  const startTime = Date.now();

  // Scroll to top first
  await frame.evaluate((selector) => {
    const container = document.querySelector(selector) ||
                      document.querySelector(".ms-DetailsList") ||
                      document.querySelector(".ms-List") ||
                      document.scrollingElement ||
                      document.documentElement;
    container.scrollTop = 0;
    window.scrollTo(0, 0);
  }, SELECTORS.SCROLLABLE_CONTAINER);
  await page.waitForTimeout(SCRAPER_CONFIG.PASS_DELAY_MS);

  while (true) {
    // Timeout guard
    if ((Date.now() - startTime) > SCRAPER_CONFIG.TOTAL_TIMEOUT_MS) {
      console.warn("Timeout reached; returning accumulated rows.");
      break;
    }

    // Get currently visible rows
    const visibleRows = await frame.locator(SELECTORS.DETAILS_ROW).all();
    let newRowsFound = false;

    for (const row of visibleRows) {
      // Get row index
      let idx = await row.getAttribute('data-item-index');
      if (!idx) {
        const rowId = await row.getAttribute('id');
        const m = (rowId || "").match(/-(\d+)$/);
        if (m) idx = m[1];
      }
      
      const indexNum = idx ? parseInt(idx, 10) : NaN;
      if (!Number.isFinite(indexNum) || seenIndices.has(indexNum)) continue;

      // Mark as seen
      seenIndices.add(indexNum);
      newRowsFound = true;

      // Extract basic row data
      const fieldContainer = row.locator(SELECTORS.DETAILS_ROW_FIELDS);
      const cells = await fieldContainer.locator(SELECTORS.DETAILS_ROW_CELL).all();
      
      const obj = {};
      for (let i = 0; i < cells.length; i++) {
        const key = await cells[i].getAttribute("data-automation-key") || `col${i}`;
        const value = await cells[i].innerText();
        obj[key] = value.trim();
      }

      // Extract details by clicking the row (with retry logic for empty descriptions)
      let details = { url: '', description: '', overview: '' };
      
      if (extractDetails) {
        for (let attempt = 1; attempt <= SCRAPER_CONFIG.MAX_RETRY_ATTEMPTS; attempt++) {
          details = await extractRowDetails(page, frame, indexNum, obj.title || '');
          
          if (details.description) {
            // Description found, break out of retry loop
            break;
          }
          
          if (attempt < SCRAPER_CONFIG.MAX_RETRY_ATTEMPTS) {
            console.warn(`Empty description for row ${indexNum}, retrying (attempt ${attempt}/${SCRAPER_CONFIG.MAX_RETRY_ATTEMPTS})...`);
          } else {
            console.warn(`Empty description for row ${indexNum} after ${SCRAPER_CONFIG.MAX_RETRY_ATTEMPTS} attempts`);
          }
        }
      }
      
      obj.url = details.url;
      obj.description = details.description;
      obj.overview = details.overview;

      processedRows.push(obj);
    }

    if (newRowsFound) {
      idlePasses = 0;
    } else {
      idlePasses++;
    }

    // Check if at bottom
    const scrollInfo = await frame.evaluate((selector) => {
      const container = document.querySelector(selector) ||
                        document.querySelector(".ms-DetailsList") ||
                        document.querySelector(".ms-List") ||
                        document.scrollingElement ||
                        document.documentElement;
      const maxScroll = Math.max(0, container.scrollHeight - container.clientHeight);
      const atBottom = container.scrollTop >= maxScroll - 2;
      return { atBottom, currentScroll: container.scrollTop, maxScroll };
    }, SELECTORS.SCROLLABLE_CONTAINER);

    if (scrollInfo.atBottom && idlePasses >= SCRAPER_CONFIG.MAX_IDLE_PASSES) {
      console.log("Reached bottom and stabilized.");
      break;
    }

    // Scroll down
    await frame.evaluate(({ selector, step }) => {
      const container = document.querySelector(selector) ||
                        document.querySelector(".ms-DetailsList") ||
                        document.querySelector(".ms-List") ||
                        document.scrollingElement ||
                        document.documentElement;
      const maxScroll = Math.max(0, container.scrollHeight - container.clientHeight);
      container.scrollTop = Math.min(container.scrollTop + step, maxScroll);
      window.scrollTo(0, container.scrollTop);
    }, { selector: SELECTORS.SCROLLABLE_CONTAINER, step: SCRAPER_CONFIG.SCROLL_STEP_PX });

    await page.waitForTimeout(SCRAPER_CONFIG.PASS_DELAY_MS);
  }

  console.log(`✅ Collected and processed ${processedRows.length} unique rows with details`);
  return processedRows;
}

module.exports = {
  waitForSplashScreen,
  clickTab,
  setDateRangeFilter,
  extractRowDetails,
  scrapeDetailsList,
};
