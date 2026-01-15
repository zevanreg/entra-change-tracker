const msal = require("@azure/msal-node");
const fs = require("fs");
const path = require("path");

// Node 18+ required (global fetch)
if (typeof fetch !== "function") {
  throw new Error("Node 18+ is required (global fetch not found).");
}

const validDateFilters = ["Last 1 month", "Last 3 months", "Last 6 months", "Last 1 year"];

// Global configuration variables
let sharepointConfig = null;
let dateFilter = null;
let accessToken = null;

/**
 * Acquire an access token using device code flow for Microsoft Graph
 * @param {object} config - Configuration object with clientId and tenantId
 * @returns {Promise<string>} Access token
 */
async function getGraphAccessTokenWithDeviceCode(config) {
  const msalConfig = {
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
    },
  };

  const pca = new msal.PublicClientApplication(msalConfig);

  const deviceCodeRequest = {
    deviceCodeCallback: (response) => {
      console.log("\n🔐 Device Code Authentication Required");
      console.log("=".repeat(60));
      console.log(response.message);
      console.log("=".repeat(60));
    },
    scopes: ["https://graph.microsoft.com/.default"],
  };

  try {
    const response = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
    return response.accessToken;
  } catch (error) {
    console.error("❌ Error acquiring token:", error.message);
    throw error;
  }
}

/**
 * Initialize configuration from config.json and acquire access token
 * @returns {Promise<void>}
 */
async function initializeConfiguration() {
  try {
    const configPath = path.join(__dirname, "..", "config.json");
    if (!fs.existsSync(configPath)) {
      console.log("⚠️ config.json not found - data will only be saved locally");
      console.log("📅 No date filter specified, showing all results");
      return;
    }

    const configContent = fs.readFileSync(configPath, "utf-8");
    const config = JSON.parse(configContent);

    // Validate and set date filter
    if (config.dateFilter) {
      if (validDateFilters.includes(config.dateFilter)) {
        dateFilter = config.dateFilter;
        console.log(`📅 Using date filter from config: ${dateFilter}`);
      } else {
        console.error(`❌ Invalid date filter in config: "${config.dateFilter}"`);
        console.error(`   Valid options: ${validDateFilters.join(", ")}`);
        process.exit(1);
      }
    } else {
      console.log("📅 No date filter specified in config, showing all results");
    }

    if (config.siteUrl && config.clientId && config.tenantId) {
      console.log("🔑 Acquiring Graph access token via device code flow...");
      accessToken = await getGraphAccessTokenWithDeviceCode(config);
      console.log("✅ Access token acquired successfully");

      sharepointConfig = config;
      console.log("✅ SharePoint/Graph configuration loaded from config.json");
    } else {
      console.log("⚠️ config.json incomplete (siteUrl/clientId/tenantId missing) - data saved locally only");
    }
  } catch (err) {
    console.error("⚠️ Error loading SharePoint config:", err.message);
    sharepointConfig = null;
  }
}

/**
 * Get the current configuration
 * @returns {{sharepointConfig: object|null, dateFilter: string|null, accessToken: string|null}}
 */
function getConfiguration() {
  return {
    sharepointConfig,
    dateFilter,
    accessToken,
  };
}

module.exports = {
  initializeConfiguration,
  getConfiguration,
  validDateFilters,
};
