/**
 * Entra Change Tracker - Main Script
 * Orchestrates scraping and data insertion into SharePoint
 */

const fs = require("fs");
const path = require("path");

// Import modules
const { initializeConfiguration, getConfiguration } = require("./lib/auth");
const { insertIntoSharePointList, resetCaches } = require("./lib/sharepoint");
const { scrapeAllSources } = require("./lib/scraper");

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

    const { config, accessToken } = getConfiguration();
    const dateFilter = config.browserScraping.dateFilter;

    // Generate timestamp for file names
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);

    // Scrape data from all sources
    const { roadmap, changeAnnouncements, whatsNew } = await scrapeAllSources(dateFilter);

    // Process Roadmap data
    if (roadmap) {
      if (config.browserScraping.roadmap.saveToFile) {
        saveToFile('roadmap', roadmap, timestamp);
      }
      
      const roadmapListName = config.browserScraping.roadmap.sharepointList.name;
      await insertIntoSharePointList(roadmapListName, roadmap, accessToken, config);
    }

    // Process Change Announcements data
    if (changeAnnouncements) {
      if (config.browserScraping.changeAnnouncements.saveToFile) {
        saveToFile('change-announcements', changeAnnouncements, timestamp);
      }
      
      const changeAnnouncementsListName = config.browserScraping.changeAnnouncements.sharepointList.name;
      await insertIntoSharePointList(changeAnnouncementsListName, changeAnnouncements, accessToken, config);
    }

    // Process What's New data
    if (whatsNew) {
      if (config.httpScraping.saveToFile) {
        saveToFile('whats-new', whatsNew, timestamp);
      }
      
      const whatsNewListName = config.httpScraping.sharepointList.whatsNew.name;
      await insertIntoSharePointList(whatsNewListName, whatsNew, accessToken, config);
    }

    console.log('✅ Script completed successfully.');
  } catch (error) {
    console.error('❌ Error during execution:', error.message);
    process.exit(1);
  }
})();
