# Entra Change Tracker - Library Modules

This directory contains reusable modules for the Entra Change Tracker application.

## Module Overview

### auth.js
Handles authentication and configuration management.

**Functions:**
- `initializeConfiguration()` - Loads config.json and acquires access token via device code flow
- `getConfiguration()` - Returns current configuration (config, dateFilter, accessToken)

**Exports:**
- `initializeConfiguration`
- `getConfiguration`
- `validDateFilters` - Array of valid date filter options

### sharepoint.js
Manages SharePoint operations using Microsoft Graph API.

**Functions:**
- `insertIntoSharePointList(listName, data, accessToken, config)` - Inserts scraped data into a SharePoint list
- `resetCaches()` - Resets internal caches for site ID and list IDs

**Internal Functions:**
- `graphFetch(token, method, graphUrl, body)` - Makes Graph API calls
- `getSiteIdFromSiteUrl(token, siteUrl)` - Resolves SharePoint site ID
- `getListIdByTitle(token, siteId, listTitle)` - Gets list ID by title
- `createListItem(token, siteId, listId, fields)` - Creates a list item
- `mapScrapedItemToSharePointFields(listName, item, config)` - Maps scraped data to SharePoint fields

**Exports:**
- `insertIntoSharePointList`
- `resetCaches`

### scraper.js
Handles all browser automation and data extraction from Entra portal.

**Functions:**
- `initializeBrowser()` - Initializes browser, navigates to Entra portal, and returns context, page, and frame
- `scrapeTab(page, frame, tabName, dateFilter)` - Scrapes data from a specific tab
- `scrapeEntraPortal(dateFilter)` - Main scraping function that extracts both Roadmap and Change Announcements

**Exports:**
- `initializeBrowser`
- `scrapeTab`
- `scrapeEntraPortal`

### browser-helpers.js
Low-level browser automation utilities for Playwright.

**Functions:**
- `waitForSplashScreen(frame, timeout)` - Waits for Entra splash screen to disappear
- `clickTab(frame, tabName)` - Clicks a tab within a frame
- `setDateRangeFilter(frame, filterOption)` - Sets the date range filter
- `extractRowDetails(page, frame, rowIndex)` - Extracts details from a row by clicking it
- `scrapeDetailsList(page, frame)` - Scrapes all rows from a Fluent UI DetailsList

**Exports:**
- `waitForSplashScreen`
- `clickTab`
- `setDateRangeFilter`
- `extractRowDetails`
- `scrapeDetailsList`

## Usage Example

```javascript
const { initializeConfiguration, getConfiguration } = require("./lib/auth");
const { insertIntoSharePointList, resetCaches } = require("./lib/sharepoint");
const { scrapeDetailsList } = require("./lib/browser-helpers");

// Initialize
await initializeConfiguration();
resetCaches();

const { config, accessToken } = getConfiguration();

// Scrape data
const data = await scrapeDetailsList(page, frame);

// Insert into SharePoint
await insertIntoSharePointList("MyList", data, accessToken, config);
```

## Dependencies

- `@azure/msal-node` - For device code authentication
- `playwright` - For browser automation
- Node.js 18+ (for global fetch API)
