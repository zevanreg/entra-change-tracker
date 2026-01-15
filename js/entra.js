/**
 * Entra Change Tracker - Main Script
 * Orchestrates scraping and data insertion into SharePoint
 */

const fs = require("fs");
const path = require("path");

// Import modules
const { initializeConfiguration, getConfiguration } = require("./lib/auth");
const { insertIntoSharePointList, resetCaches } = require("./lib/sharepoint");
const { scrapeEntraPortal } = require("./lib/scraper");

/**
 * Save data to JSON file
 * @param {string} filename - Base filename
 * @param {Array<Object>} data - Data to save
 * @param {string} timestamp - Timestamp for filename
 */
function saveToFile(filename, data, timestamp) {
  const filePath = path.join(__dirname, `${filename}-${timestamp}.json`);
  fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf-8');
  console.log(`💾 Saved ${filename} to ${filePath}`);
}

/**
 * Main execution function
 */
(async () => {
  try {
    // Initialize configuration and authenticate
    await initializeConfiguration();
    resetCaches();

    const { sharepointConfig, dateFilter, accessToken } = getConfiguration();

    // Generate timestamp for file names
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);

    // Scrape data from Entra portal
    const { roadmap, changeAnnouncements } = await scrapeEntraPortal(dateFilter);

    // Process Roadmap data
    if (roadmap) {
      saveToFile('roadmap', roadmap, timestamp);
      
      const roadmapListName = sharepointConfig?.lists?.roadmap || 'EntraRoadmapItems';
      await insertIntoSharePointList(roadmapListName, roadmap, accessToken, sharepointConfig);
    }

    // Process Change Announcements data
    if (changeAnnouncements) {
      saveToFile('change-announcements', changeAnnouncements, timestamp);
      
      const changeAnnouncementsListName = sharepointConfig?.lists?.changeAnnouncements || 'EntraChangeAnnouncements';
      await insertIntoSharePointList(changeAnnouncementsListName, changeAnnouncements, accessToken, sharepointConfig);
    }

    console.log('✅ Script completed successfully.');
  } catch (error) {
    console.error('❌ Error during execution:', error.message);
    process.exit(1);
  }
})();
