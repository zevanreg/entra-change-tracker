const msal = require("@azure/msal-node");
const { 
  DataProtectionScope, 
  Environment, 
  PersistenceCreator, 
  PersistenceCachePlugin 
} = require("@azure/msal-node-extensions");
const fs = require("fs");
const path = require("path");

// Node 18+ required (global fetch)
if (typeof fetch !== "function") {
  throw new Error("Node 18+ is required (global fetch not found).");
}

const validDateFilters = ["Last 1 month", "Last 3 months", "Last 6 months", "Last 1 year"];

// Global configuration variables
let config = null;
let dateFilter = null;
let accessToken = null;

/**
 * Create and configure the MSAL cache plugin for token persistence
 * @returns {Promise<PersistenceCachePlugin>} Configured cache plugin
 */
async function createCachePlugin() {
  const cachePath = path.join(__dirname, "..", ".token-cache.json");
  
  const persistenceConfiguration = {
    cachePath,
    dataProtectionScope: DataProtectionScope.CurrentUser,
    serviceName: "entra-change-tracker",
    accountName: "msal-token-cache",
    usePlaintextFileOnLinux: true,
  };

  const persistence = await PersistenceCreator.createPersistence(persistenceConfiguration);
  return new PersistenceCachePlugin(persistence);
}

/**
 * Acquire an access token using device code flow for Microsoft Graph
 * @param {object} config - Configuration object with clientId and tenantId
 * @returns {Promise<string>} Access token
 */
async function getGraphAccessTokenWithDeviceCode(config) {
  const cachePlugin = await createCachePlugin();
  
  const msalConfig = {
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
    },
    cache: {
      cachePlugin,
    },
  };

  const pca = new msal.PublicClientApplication(msalConfig);
  
  // Try silent token acquisition first
  const accounts = await pca.getTokenCache().getAllAccounts();
  
  if (accounts.length > 0) {
    try {
      console.log("🔄 Attempting to use cached token...");
      const silentRequest = {
        account: accounts[0],
        scopes: ["https://graph.microsoft.com/.default"],
      };
      const response = await pca.acquireTokenSilent(silentRequest);
      console.log("✅ Using cached token");
      return response.accessToken;
    } catch (error) {
      console.log("⚠️ Cached token invalid or expired, requesting new token...");
    }
  }

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
    const loadedConfig = JSON.parse(configContent);

    // Validate and set date filter
    if (loadedConfig.dateFilter) {
      if (validDateFilters.includes(loadedConfig.dateFilter)) {
        dateFilter = loadedConfig.dateFilter;
        console.log(`📅 Using date filter from config: ${dateFilter}`);
      } else {
        console.error(`❌ Invalid date filter in config: "${loadedConfig.dateFilter}"`);
        console.error(`   Valid options: ${validDateFilters.join(", ")}`);
        process.exit(1);
      }
    } else {
      console.log("📅 No date filter specified in config, showing all results");
    }

    if (loadedConfig.siteUrl && loadedConfig.clientId && loadedConfig.tenantId) {
      console.log("🔑 Acquiring Graph access token via device code flow...");
      accessToken = await getGraphAccessTokenWithDeviceCode(loadedConfig);
      console.log("✅ Access token acquired successfully");

      config = loadedConfig;
      console.log("✅ SharePoint/Graph configuration loaded from config.json");
    } else {
      console.log("⚠️ config.json incomplete (siteUrl/clientId/tenantId missing) - data saved locally only");
    }
  } catch (err) {
    console.error("⚠️ Error loading SharePoint config:", err.message);
    config = null;
  }
}

/**
 * Get the current configuration
 * @returns {{config: object|null, dateFilter: string|null, accessToken: string|null}}
 */
function getConfiguration() {
  return {
    config,
    dateFilter,
    accessToken,
  };
}

module.exports = {
  initializeConfiguration,
  getConfiguration,
  validDateFilters,
};
