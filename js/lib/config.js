/**
 * Configuration management for Entra Change Tracker
 * Handles loading and validation of config.json
 */

const fs = require("fs");
const path = require("path");

const VALID_DATE_FILTERS = ["Last 1 month", "Last 3 months", "Last 6 months", "Last 1 year"];

// Global configuration state
let _config = null;
let _dateFilter = null;

/**
 * Load and validate configuration from config.json
 * @returns {void}
 */
function loadConfiguration() {
  try {
    const configPath = path.join(__dirname, "..", "config.json");
    
    if (!fs.existsSync(configPath)) {
      console.log("⚠️ config.json not found - data will only be saved locally");
      console.log("📅 No date filter specified, showing all results");
      return;
    }

    const configContent = fs.readFileSync(configPath, "utf-8");
    _config = JSON.parse(configContent);

    // Validate and set date filter
    const configDateFilter = _config.browserScraping?.dateFilter;
    if (configDateFilter) {
      if (VALID_DATE_FILTERS.includes(configDateFilter)) {
        _dateFilter = configDateFilter;
        console.log(`📅 Using date filter from config: ${_dateFilter}`);
      } else {
        console.error(`❌ Invalid date filter in config: "${configDateFilter}"`);
        console.error(`   Valid options: ${VALID_DATE_FILTERS.join(", ")}`);
        process.exit(1);
      }
    } else {
      console.log("📅 No date filter specified in config, showing all results");
    }
  } catch (err) {
    console.error("❌ Error loading config.json:", err.message);
    process.exit(1);
  }
}

/**
 * Get the loaded configuration
 * @returns {object|null} The configuration object
 */
function getConfig() {
  return _config;
}

/**
 * Get the date filter from configuration
 * @returns {string|null} The date filter string
 */
function getDateFilter() {
  return _dateFilter;
}

/**
 * Check if configuration is loaded
 * @returns {boolean} True if configuration is loaded
 */
function isConfigLoaded() {
  return _config !== null;
}

module.exports = {
  loadConfiguration,
  getConfig,
  getDateFilter,
  isConfigLoaded,
  VALID_DATE_FILTERS,
};
