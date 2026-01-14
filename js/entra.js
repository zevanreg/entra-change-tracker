// test-entra-roadmap.js
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const { sp } = require('@pnp/sp');
const msal = require('@azure/msal-node');

const url = 'https://entra.microsoft.com/#blade/Microsoft_AAD_IAM/ChangeManagementHubList.ReactView';

// Valid date filter options
const validDateFilters = ['Last 1 month', 'Last 3 months', 'Last 6 months', 'Last 1 year'];

/**
 * Acquire an access token using device code flow
 * @param {object} config - Configuration object with clientId and tenantId
 * @returns {Promise<string>} Access token
 */
async function getAccessTokenWithDeviceCode(config) {
  const msalConfig = {
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
    },
  };

  const pca = new msal.PublicClientApplication(msalConfig);
  const deviceCodeRequest = {
    deviceCodeCallback: (response) => {
      console.log('\n🔐 Device Code Authentication Required');
      console.log('=' .repeat(60));
      console.log(response.message);
      console.log('=' .repeat(60));
    },
    scopes: [`${config.siteUrl}/.default`],
  };

  try {
    const response = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
    return response.accessToken;
  } catch (error) {
    console.error('❌ Error acquiring token:', error.message);
    throw error;
  }
}

// Load SharePoint configuration from file
let sharepointConfig = null;
let dateFilter = null;
let accessToken = null;

try {
  const configPath = path.join(__dirname, 'config.json');
  if (fs.existsSync(configPath)) {
    const configContent = fs.readFileSync(configPath, 'utf-8');
    const config = JSON.parse(configContent);
    
    // Validate and set date filter
    if (config.dateFilter) {
      if (validDateFilters.includes(config.dateFilter)) {
        dateFilter = config.dateFilter;
        console.log(`📅 Using date filter from config: ${dateFilter}`);
      } else {
        console.error(`❌ Invalid date filter in config: "${config.dateFilter}"`);
        console.error(`   Valid options: ${validDateFilters.join(', ')}`);
        process.exit(1);
      }
    } else {
      console.log('📅 No date filter specified in config, showing all results');
    }
    
    // Configure PnPjs if all required fields are present
    if (config.siteUrl && config.clientId && config.tenantId) {
      console.log('🔑 Acquiring access token via device code flow...');
      accessToken = await getAccessTokenWithDeviceCode(config);
      console.log('✅ Access token acquired successfully');
      
      sp.setup({
        sp: {
          fetchClientFactory: () => {
            return () => {
              return fetch(arguments[0], {
                ...arguments[1],
                headers: {
                  ...arguments[1]?.headers,
                  'Authorization': `Bearer ${accessToken}`,
                },
              });
            };
          },
        },
      });
      console.log('✅ SharePoint configuration loaded from config.json');
      sharepointConfig = config;
    } else {
      console.log('⚠️ SharePoint config file incomplete - data will only be saved locally');
    }
  } else {
    console.log('⚠️ config.json not found - data will only be saved locally');
    console.log('📅 No date filter specified, showing all results');
  }
} catch (err) {
  console.error('⚠️ Error loading SharePoint config:', err.message);
  sharepointConfig = null;
}

/**
 * Insert extracted data into a SharePoint list
 * @param {string} listName - The name of the SharePoint list
 * @param {object[]} data - The array of data items to insert
 */
async function insertIntoSharePointList(listName, data) {
  if (!sharepointConfig) {
    console.log(`⏭️ Skipping SharePoint insertion for ${listName} (not configured)`);
    return;
  }

  try {
    console.log(`📤 Inserting ${data.length} items into SharePoint list: ${listName}`);
    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < data.length; i++) {
      try {
        await sp.web.lists.getByTitle(listName).items.add(data[i]);
        successCount++;
        if ((i + 1) % 10 === 0) {
          console.log(`   Progress: ${i + 1}/${data.length} items inserted`);
        }
      } catch (itemErr) {
        errorCount++;
        console.error(`   Error inserting item ${i + 1}:`, itemErr.message);
      }
    }

    console.log(`✅ Inserted ${successCount} items into ${listName} (${errorCount} errors)`);
  } catch (err) {
    console.error(`❌ Error inserting into ${listName}:`, err.message);
  }
}

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
 * @returns {Promise<{url: string, description: string}>} The extracted details
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
      return { url: '', description: '' };
    }
    
    // Wait for progress dots to disappear
    try {
      const progressDots = detailsFrame.locator('div.fxs-progress-dots');
      await progressDots.waitFor({ state: 'hidden', timeout: 15000 });
    } catch (err) {
      // Progress dots might not exist or already hidden
    }
    
    // Extract URL from "Next steps" section
    let url = '';
    try {
      // Verify frame is still attached
      if (!detailsFrame.isDetached()) {
        const nextStepsH3 = detailsFrame.locator('h3').filter({ hasText: 'Next steps' });
        
        if (await nextStepsH3.count() > 0) {
          // Navigate to parent and find link
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
    
    // Extract description from "Here's what you will see in this release:" section
    let description = '';
    try {
      // Verify frame is still attached
      if (!detailsFrame.isDetached()) {
        const releaseH3 = detailsFrame.locator('h3').filter({ hasText: "Here's what you will see in this release:" });
        
        if (await releaseH3.count() > 0) {
          // Navigate to parent and find paragraph
          const parent = releaseH3.locator('..');
          const paragraphs = parent.locator('p');
          
          if (await paragraphs.count() > 0) {
            description = await paragraphs.first().innerText() || '';
          }
        }
      }
    } catch (err) {
      console.warn(`Could not extract description for row ${rowIndex}:`, err.message);
    }
    
    // // Close the details pane - look for Azure blade close button
    // try {
    //   // Try multiple selectors for the close button
    //   const closeSelectors = [
    //     'button.fxs-blade-close',
    //     'button[class*="fxs-blade-close"]',
    //     'button[aria-label="Close"]',
    //     'div.fxs-blade-close-button button'
    //   ];
      
    //   let closed = false;
    //   for (const selector of closeSelectors) {
    //     const closeButton = page.locator(selector);
    //     if (await closeButton.count() > 0) {
    //       try {
    //       await closeButton.first().click({ timeout: 5000 });
    //       console.info(`Details pane closed for row ${rowIndex} using selector: ${selector}`);
    //       closed = true;
    //       break;
    //       } catch (clickErr) {
    //         console.warn(`Click failed on close button with selector ${selector}, trying next selector.`);
    //       }
    //     }
    //   }
      
      // if (!closed) {
      //   console.warn(`Close button not found for row ${rowIndex}, trying ESC key`);
        await page.keyboard.press('Escape');
      // }
      
      // await page.waitForTimeout(500);
    // } catch (err) {
    //   console.warn(`Could not close details pane for row ${rowIndex}:`, err.message);
    //   // Try ESC as last resort
    //   try {
    //     await page.keyboard.press('Escape');
    //   } catch {}
    // }
    
    return { url: url.trim(), description: description.trim() };
  } catch (err) {
    console.error(`Error extracting details for row ${rowIndex}:`, err.message);
    
    // Attempt to close any open pane
    try {
      await page.keyboard.press('Escape');
      await page.waitForTimeout(500);
    } catch {}
    
    return { url: '', description: '' };
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

  // --- Stable selectors for Fluent UI DetailsList ---
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

(async () => {
  const context = await chromium.launchPersistentContext('./edge-profile', {
    channel: 'msedge',
    headless: false, // first runs: keep visible
    args: ['--disable-blink-features=AutomationControlled'],
    viewport: { width: 1920, height: 1080 },
    deviceScaleFactor: 1,
  });
  const page = context.pages()[0];

  // 1) Navigate and wait for the blade to settle
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });

  // 2) Get the reactblade frame by element attribute (not context name)
  //const blade = page.frameLocator('iframe[name="ChangeManagementHubList.ReactView"]');
  const iframeLocator = page.locator('iframe[name="ChangeManagementHubList.ReactView"]');
  await iframeLocator.waitFor({ state: 'attached', timeout: 30000 });

  const iframeHandle = await iframeLocator.elementHandle();
  const frame = await iframeHandle.contentFrame();
  if (!frame) throw new Error('ReactView frame attached, but content not available yet.');

  // Wait for initial splash screen to disappear
  await waitForSplashScreen(frame, 60000);

  // Wait for progress dots to disappear
  try {
    const progressDots = frame.locator('div.fxs-progress-dots');
    await progressDots.waitFor({ state: 'hidden', timeout: 15000 });
  } catch (err) {
    // Progress dots might not exist or already hidden
  }

  // Click the Roadmap tab
  const roadmapClicked = await clickTab(frame, /^Roadmap$/i);

  if (!roadmapClicked) {
    console.error('❌ Could not locate the Roadmap tab/menu. Opening inspector...');
    await context.close();
    return;
  }

  // Set the date range filter only if specified
  if (dateFilter) {
    const filterSet = await setDateRangeFilter(frame, dateFilter);
    if (!filterSet) {
      console.warn('⚠️ Could not set date range filter, continuing anyway...');
    }
  }

  const roadmapRows = await scrapeDetailsList(page, frame);

  console.log(roadmapRows.slice(0, 5));
  console.log(`✅ Extracted ${roadmapRows.length} roadmap items`);

  // Save roadmap data
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
  const roadmapFile = path.join(__dirname, `roadmap-${timestamp}.json`);
  fs.writeFileSync(roadmapFile, JSON.stringify(roadmapRows, null, 2), 'utf-8');
  console.log(`💾 Saved roadmap to ${roadmapFile}`);

  // Insert roadmap data into SharePoint list
  const roadmapListName = sharepointConfig?.lists?.roadmap || 'EntraRoadmapItems';
  await insertIntoSharePointList(roadmapListName, roadmapRows);

  // Click the Change announcements tab
  const changesClicked = await clickTab(frame, /^Change announcements$/i);

  if (!changesClicked) {
    console.error('❌ Could not locate the Change announcements tab/menu. Opening inspector...');
    await context.close();
    return;
  }

  // Set the date range filter only if specified
  if (dateFilter) {
    const filterSet = await setDateRangeFilter(frame, dateFilter);
    if (!filterSet) {
      console.warn('⚠️ Could not set date range filter, continuing anyway...');
    }
  }

  const changeAnnouncementRows = await scrapeDetailsList(page, frame);

  console.log(changeAnnouncementRows.slice(0, 5));
  console.log(`✅ Extracted ${changeAnnouncementRows.length} change announcements`);

  // Save change announcements data
  const changesFile = path.join(__dirname, `change-announcements-${timestamp}.json`);
  fs.writeFileSync(changesFile, JSON.stringify(changeAnnouncementRows, null, 2), 'utf-8');
  console.log(`💾 Saved change announcements to ${changesFile}`);

  // Insert change announcements into SharePoint list
  const changeAnnouncementsListName = sharepointConfig?.lists?.changeAnnouncements || 'EntraChangeAnnouncements';
  await insertIntoSharePointList(changeAnnouncementsListName, changeAnnouncementRows);

  await context.close();
})();
