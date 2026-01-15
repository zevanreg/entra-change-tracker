/**
 * Browser automation helper functions for Entra portal scraping
 */

/**
 * Waits for the Entra splash screen to disappear
 * @param {import('playwright').Frame} frame - The frame to check
 * @param {number} timeout - Maximum wait time in ms
 */
async function waitForSplashScreen(frame, timeout = 30000) {
  try {
    const splashScreen = frame.locator('.fxs-splashscreen');
    await splashScreen.waitFor({ state: 'hidden', timeout });
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
  try {
    // Wait for splash screen to disappear first
    await waitForSplashScreen(frame);
    
    const roleTab = frame.getByRole('tab', { name: tabName });
    const count = await roleTab.count();
    if (count > 0) {
      // Wait for tab to be visible and stable
      await roleTab.first().waitFor({ state: 'visible', timeout: 10000 });
      
      try {
        // Try normal click first
        await roleTab.first().click({ timeout: 5000 });
      } catch (clickErr) {
        // If intercepted, force click
        console.log(`Normal click failed, forcing click on tab "${tabName}"`);
        await roleTab.first().click({ force: true, timeout: 5000 });
      }
      
      // Wait a moment for tab content to load
      await frame.waitForTimeout(1000);
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
  try {
    // Click the filter button - find button inside div with data-selection-index='1'
    const filterButton = frame.locator('div[data-selection-index="1"] button').first();
    await filterButton.waitFor({ state: 'visible', timeout: 10000 });
    await filterButton.click({ timeout: 5000 });
    console.log('Filter button clicked');
    
    // Wait for the filter menu to appear
    await frame.waitForTimeout(500);
    
    // Find and click the radio input with the matching label
    const radioLabel = frame.locator(`span.ms-ChoiceFieldLabel:has-text("${filterOption}")`);
    await radioLabel.waitFor({ state: 'visible', timeout: 5000 });
    
    // Click the label to select the radio button
    await radioLabel.click({ timeout: 5000 });
    console.log(`Selected filter option: ${filterOption}`);
    
    // Wait a moment for the selection to register
    await frame.waitForTimeout(500);
    
    // Click the Apply button
    const applyButton = frame.locator('button:has-text("Apply")');
    if (await applyButton.count() > 0 && !await applyButton.isDisabled()) {
      await applyButton.click({ timeout: 5000 });
      console.log('Apply button clicked');
      await frame.waitForTimeout(1000);
    }
    
    return true;
  } catch (err) {
    console.error(`Error setting date range filter to "${filterOption}":`, err.message);
    return false;
  }
}

/**
 * Extracts details from a row by clicking it and scraping the details pane
 * @param {import('playwright').Page} page - The main page
 * @param {import('playwright').Frame} frame - The frame containing the row
 * @param {number} rowIndex - The index of the row to click
 * @returns {Promise<{url: string, description: string, overview: string}>} The extracted details
 */
async function extractRowDetails(page, frame, rowIndex) {
  try {
    // Click the row to open details pane
    const row = frame.locator(`div[data-automationid='DetailsRow'][data-item-index='${rowIndex}']`).first();
    await row.click({ timeout: 5000 });
    
    // Wait for details iframe to appear
    const detailsIframeLocator = page.locator('iframe[name="ChangeManagementHubEntityDetailsPane.ReactView"]').last();
    await detailsIframeLocator.waitFor({ state: 'attached', timeout: 10000 });
    
    const detailsIframeHandle = await detailsIframeLocator.elementHandle();
    const detailsFrame = await detailsIframeHandle.contentFrame();
    
    if (!detailsFrame) {
      console.warn(`Details frame not available for row ${rowIndex}`);
      return { url: '', description: '', overview: '' };
    }
    
    // Wait for progress dots to disappear
    try {
      const progressDots = detailsFrame.locator('div.fxs-progress-dots');
      await progressDots.waitFor({ state: 'hidden', timeout: 15000 });
    } catch (err) {
      // Progress dots might not exist or already hidden
    }
    
    // Extract overview from "Overview" section
    let overview = '';
    try {
      if (!detailsFrame.isDetached()) {
        const overviewH3 = detailsFrame.locator('h3').filter({ hasText: 'Overview' });
        
        if (await overviewH3.count() > 0) {
          const parent = overviewH3.locator('..');
          const spans = parent.locator('span');
          
          if (await spans.count() > 0) {
            overview = await spans.first().innerText() || '';
          }
        }
      }
    } catch (err) {
      console.warn(`Could not extract overview for row ${rowIndex}:`, err.message);
    }
    
    // Extract URL from "Next steps" section
    let url = '';
    try {
      if (!detailsFrame.isDetached()) {
        const nextStepsH3 = detailsFrame.locator('h3').filter({ hasText: 'Next steps' });
        
        if (await nextStepsH3.count() > 0) {
          const parent = nextStepsH3.locator('..');
          const links = parent.locator('a');
          
          if (await links.count() > 0) {
            url = await links.first().getAttribute('href') || '';
          }
        }
      }
    } catch (err) {
      console.warn(`Could not extract URL for row ${rowIndex}:`, err.message);
    }
    
    // Extract description from "What is changing" section (for change announcements)
    // or "Here's what you will see in this release:" section (for roadmap)
    let description = '';
    try {
      if (!detailsFrame.isDetached()) {
        // Try "What is changing" first (change announcements)
        let descH3 = detailsFrame.locator('h3').filter({ hasText: 'What is changing' });
        
        if (await descH3.count() > 0) {
          const parent = descH3.locator('..');
          const spans = parent.locator('span');
          
          if (await spans.count() > 0) {
            description = await spans.first().innerText() || '';
          }
        } else {
          // Fall back to "Here's what you will see in this release:" (roadmap)
          descH3 = detailsFrame.locator('h3').filter({ hasText: "Here's what you will see in this release:" });
          
          if (await descH3.count() > 0) {
            const parent = descH3.locator('..');
            const paragraphs = parent.locator('p');
            
            if (await paragraphs.count() > 0) {
              description = await paragraphs.first().innerText() || '';
            }
          }
        }
      }
    } catch (err) {
      console.warn(`Could not extract description for row ${rowIndex}:`, err.message);
    }
    
    // Close details pane with ESC key
    await page.keyboard.press('Escape');
    
    return { url: url.trim(), description: description.trim(), overview: overview.trim() };
  } catch (err) {
    console.error(`Error extracting details for row ${rowIndex}:`, err.message);
    
    // Attempt to close any open pane
    try {
      await page.keyboard.press('Escape');
      await page.waitForTimeout(500);
    } catch {}
    
    return { url: '', description: '', overview: '' };
  }
}

/**
 * Scrapes all rows from a Fluent UI DetailsList with virtualized scrolling
 * @param {import('playwright').Page} page - The main page
 * @param {import('playwright').Frame} frame - The frame containing the DetailsList
 * @returns {Promise<Array<Object>>} Array of row objects with mapped field names and details
 */
async function scrapeDetailsList(page, frame) {
  // Field name mapping
  const fieldMap = {
    'title': 'title',
    'changeEntityCategory': 'category',
    'changeEntityService': 'service',
    'changeEntityDeliveryStage': 'releaseType',
    'publishStartDateTime': 'releaseDate',
    'changeEntityState': 'state'
  };

  // Stable selectors for Fluent UI DetailsList
  const ROW_SELECTOR = "div[data-automationid='DetailsRow']";
  const FIELD_CONTAINER_SELECTOR = "div[data-automationid='DetailsRowFields']";
  const CELL_SELECTOR = "div[data-automationid='DetailsRowCell']";
  const INDEX_ATTR = "data-item-index";

  // Accumulator for processed rows
  const processedRows = [];
  const seenIndices = new Set();

  // Timing & scroll config
  const STEP_PX = 0.85 * 1080; // Based on viewport height
  const PASS_DELAY_MS = 300;
  const MAX_IDLE_PASSES = 6;
  const TOTAL_TIMEOUT_MS = 180000;

  let idlePasses = 0;
  const startTime = Date.now();

  // Scroll to top first
  await frame.evaluate(() => {
    const container = document.querySelector("div[data-is-scrollable='true']") ||
                      document.querySelector(".ms-DetailsList") ||
                      document.querySelector(".ms-List") ||
                      document.scrollingElement ||
                      document.documentElement;
    container.scrollTop = 0;
    window.scrollTo(0, 0);
  });
  await page.waitForTimeout(PASS_DELAY_MS);

  while (true) {
    // Timeout guard
    if ((Date.now() - startTime) > TOTAL_TIMEOUT_MS) {
      console.warn("Timeout reached; returning accumulated rows.");
      break;
    }

    // Get currently visible rows
    const visibleRows = await frame.locator(ROW_SELECTOR).all();
    let newRowsFound = false;

    for (const row of visibleRows) {
      // Get row index
      let idx = await row.getAttribute(INDEX_ATTR);
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
      const fieldContainer = row.locator(FIELD_CONTAINER_SELECTOR);
      const cells = await fieldContainer.locator(CELL_SELECTOR).all();
      
      const obj = {};
      for (let i = 0; i < cells.length; i++) {
        const key = await cells[i].getAttribute("data-automation-key") || `col${i}`;
        const value = await cells[i].innerText();
        obj[key] = value.trim();
      }

      // Map field names
      const mapped = {};
      Object.keys(obj).forEach(key => {
        const newKey = fieldMap[key] || key;
        mapped[newKey] = obj[key];
      });

      console.log(`Processing row ${processedRows.length + 1}: ${mapped.title || 'Untitled'}`);

      // Extract details by clicking the row
      const details = await extractRowDetails(page, frame, indexNum);
      mapped.url = details.url;
      mapped.description = details.description;
      mapped.overview = details.overview;

      processedRows.push(mapped);
    }

    if (newRowsFound) {
      idlePasses = 0;
    } else {
      idlePasses++;
    }

    // Check if at bottom
    const scrollInfo = await frame.evaluate(() => {
      const container = document.querySelector("div[data-is-scrollable='true']") ||
                        document.querySelector(".ms-DetailsList") ||
                        document.querySelector(".ms-List") ||
                        document.scrollingElement ||
                        document.documentElement;
      const maxScroll = Math.max(0, container.scrollHeight - container.clientHeight);
      const atBottom = container.scrollTop >= maxScroll - 2;
      return { atBottom, currentScroll: container.scrollTop, maxScroll };
    });

    if (scrollInfo.atBottom && idlePasses >= MAX_IDLE_PASSES) {
      console.log("Reached bottom and stabilized.");
      break;
    }

    // Scroll down
    await frame.evaluate((step) => {
      const container = document.querySelector("div[data-is-scrollable='true']") ||
                        document.querySelector(".ms-DetailsList") ||
                        document.querySelector(".ms-List") ||
                        document.scrollingElement ||
                        document.documentElement;
      const maxScroll = Math.max(0, container.scrollHeight - container.clientHeight);
      container.scrollTop = Math.min(container.scrollTop + step, maxScroll);
      window.scrollTo(0, container.scrollTop);
    }, STEP_PX);

    await page.waitForTimeout(PASS_DELAY_MS);
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
