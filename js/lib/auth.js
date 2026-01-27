const msal = require("@azure/msal-node");
const { 
  DataProtectionScope, 
  Environment, 
  PersistenceCreator, 
  PersistenceCachePlugin 
} = require("@azure/msal-node-extensions");
const { DefaultAzureCredential } = require("@azure/identity");
const path = require("path");
const { getConfig } = require("./config");

// Node 18+ required (global fetch)
if (typeof fetch !== "function") {
  throw new Error("Node 18+ is required (global fetch not found).");
}

// Global access token
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
 * Acquire an access token using DefaultAzureCredential (IWA) for Microsoft Graph
 * Uses the currently logged-in user's credentials
 * @param {object} config - Configuration object (optional tenantId)
 * @returns {Promise<string>} Access token
 */
async function getGraphAccessTokenWithIWA(config) {
  console.log("🔄 Using DefaultAzureCredential (Integrated Windows Authentication)...");
  
  // Create credential with optional tenant ID
  const credentialOptions = {};
  if (config.tenantId) {
    credentialOptions.tenantId = config.tenantId;
  }
  
  const credential = new DefaultAzureCredential(credentialOptions);
  
  // Get token for Microsoft Graph
  const tokenResponse = await credential.getToken("https://graph.microsoft.com/.default");
  console.log("✅ Token acquired via DefaultAzureCredential");
  
  return tokenResponse.token;
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
 * Initialize authentication and acquire access token
 * Requires configuration to be loaded first via config.loadConfiguration()
 * @returns {Promise<void>}
 */
async function initializeAuthentication() {
  try {
    const config = getConfig();
    
    if (!config) {
      console.log("⚠️ No configuration loaded - authentication skipped");
      return;
    }

    // Determine authentication method
    const authMethod = (config.sharepoint?.authentication?.authMethod || "devicecode").toLowerCase();
    
    if (config.sharepoint?.siteUrl) {
      if (authMethod === "iwa" || authMethod === "default") {
        console.log("🔑 Acquiring Graph access token via DefaultAzureCredential (IWA)...");
        accessToken = await getGraphAccessTokenWithIWA(config.sharepoint.authentication);
        console.log("✅ Access token acquired successfully");
      } else if (authMethod === "devicecode") {
        const deviceCodeConfig = config.sharepoint.authentication.devicecode;
        if (!deviceCodeConfig?.clientId || !deviceCodeConfig?.tenantId) {
          console.error("❌ clientId and tenantId are required for device code flow");
          process.exit(1);
        }
        console.log("🔑 Acquiring Graph access token via device code flow...");
        accessToken = await getGraphAccessTokenWithDeviceCode(deviceCodeConfig);
        console.log("✅ Access token acquired successfully");
      } else {
        console.error(`❌ Invalid authMethod: "${authMethod}". Valid options: 'devicecode', 'iwa', 'default'`);
        process.exit(1);
      }
    } else {
      console.log("⚠️ config.json incomplete (siteUrl missing) - data saved locally only");
    }
  } catch (err) {
    console.error("⚠️ Error during authentication:", err.message);
    throw err;
  }
}

/**
 * Get the access token
 * @returns {string|null}
 */
function getAccessToken() {
  return accessToken;
}

module.exports = {
  initializeAuthentication,
  getAccessToken,
};
