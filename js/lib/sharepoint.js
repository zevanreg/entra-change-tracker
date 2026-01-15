// Graph caches
let cachedSiteId = null;
const cachedListIds = new Map(); // listTitle -> listId

/**
 * Make a Microsoft Graph API call
 * @param {string} token - Access token
 * @param {string} method - HTTP method
 * @param {string} graphUrl - Graph API URL
 * @param {object} [body] - Request body
 * @returns {Promise<object>} Response JSON
 */
async function graphFetch(token, method, graphUrl, body) {
  const res = await fetch(graphUrl, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph HTTP ${res.status} ${res.statusText}\nURL: ${graphUrl}\n${text}`);
  }

  return res.json();
}

/**
 * Get SharePoint site ID from site URL
 * @param {string} token - Access token
 * @param {string} siteUrl - SharePoint site URL
 * @returns {Promise<string>} Site ID
 */
async function getSiteIdFromSiteUrl(token, siteUrl) {
  if (cachedSiteId) return cachedSiteId;

  const u = new URL(siteUrl);
  const hostname = u.hostname; // e.g. m365j556631.sharepoint.com
  const sitePath = u.pathname.replace(/\/$/, ""); // e.g. /sites/EntraChangeTrackers

  // GET /sites/{hostname}:{server-relative-path}
  const endpoint = `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`;
  const site = await graphFetch(token, "GET", endpoint);

  if (!site?.id) throw new Error(`Could not resolve siteId for ${siteUrl}`);
  cachedSiteId = site.id;
  return cachedSiteId;
}

/**
 * Get SharePoint list ID by title
 * @param {string} token - Access token
 * @param {string} siteId - Site ID
 * @param {string} listTitle - List title
 * @returns {Promise<string>} List ID
 */
async function getListIdByTitle(token, siteId, listTitle) {
  if (cachedListIds.has(listTitle)) return cachedListIds.get(listTitle);

  const safeTitle = listTitle.replace(/'/g, "''");
  const endpoint =
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists` +
    `?$filter=displayName eq '${safeTitle}'&$select=id,displayName`;

  const result = await graphFetch(token, "GET", endpoint);
  const match = (result.value || []).find((x) => x.displayName === listTitle);

  if (!match) throw new Error(`List not found by title "${listTitle}" (check list name/spelling).`);

  cachedListIds.set(listTitle, match.id);
  return match.id;
}

/**
 * Create a list item in SharePoint
 * @param {string} token - Access token
 * @param {string} siteId - Site ID
 * @param {string} listId - List ID
 * @param {object} fields - Item fields
 * @returns {Promise<object>} Created item
 */
async function createListItem(token, siteId, listId, fields) {
  const endpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;
  return graphFetch(token, "POST", endpoint, { fields });
}

/**
 * Map scraped item to SharePoint fields
 * @param {string} listName - List name
 * @param {object} item - Scraped item
 * @param {object} config - SharePoint configuration
 * @returns {object} Mapped fields
 */
function mapScrapedItemToSharePointFields(listName, item, config) {
  // Optional per-list mapping in config.json:
  // {
  //   "fieldMappings": {
  //     "EntraRoadmapItems": {
  //       "Title": "title",
  //       "Category": "category",
  //       ...
  //     }
  //   }
  // }
  const mapping = config?.fieldMappings?.[listName];

  if (mapping) {
    const fields = {};
    for (const [spInternalName, sourceKey] of Object.entries(mapping)) {
      fields[spInternalName] = item[sourceKey];
    }
    // Ensure Title exists if possible
    if (!fields.Title && item.title) fields.Title = item.title;
    return fields;
  }

  // Default mappings based on list name
  const listNameLower = listName.toLowerCase();
  
  // Roadmap list mapping
  if (listNameLower.includes('roadmap')) {
    return {
      Title: item.title || item.Title || "",
      Category: item.category || "",
      Service: item.service || "",
      ReleaseType: item.releaseType || "",
      ReleaseDate: item.releaseDate || "",
      State: item.state || "",
      Overview: item.overview || "",
      Description: item.description || "",
      Url: item.url || "",
    };
  }
  
  // Change Announcements list mapping
  if (listNameLower.includes('change') || listNameLower.includes('announcement')) {
    return {
      Title: item.title || item.Title || "",
      Service: item.service || "",
      ChangeType: item.changeEntityChangeType || "",
      AnnouncementDate: item.announcementDateTime || "",
      TargetDate: item.targetDateTime || "",
      ActionRequired: item.isCustomerActionRequired || "",
      Tags: item.marketingThemes || "", // Use tags if available
      Overview: item.overview || "",
      Description: item.description || "",
      Url: item.url || "",
    };
  }

  // If list name doesn't match expected patterns, throw an error
  throw new Error(
    `Unknown list name: "${listName}". Expected list name to contain "roadmap" or "change"/"announcement". ` +
    `Please use proper list names or define custom fieldMappings in config.json.`
  );
}

/**
 * Insert data into a SharePoint list using Microsoft Graph API
 * @param {string} listName - The name of the SharePoint list
 * @param {object[]} data - The array of data items to insert
 * @param {string} accessToken - Access token for Graph API
 * @param {object} sharepointConfig - SharePoint configuration
 * @returns {Promise<void>}
 */
async function insertIntoSharePointList(listName, data, accessToken, sharepointConfig) {
  if (!sharepointConfig || !accessToken) {
    console.log(`⏭️ Skipping SharePoint insertion for ${listName} (not configured)`);
    return;
  }

  try {
    console.log(`📤 Inserting ${data.length} items into SharePoint list (Graph): ${listName}`);

    const siteId = await getSiteIdFromSiteUrl(accessToken, sharepointConfig.siteUrl);
    const listId = await getListIdByTitle(accessToken, siteId, listName);

    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < data.length; i++) {
      try {
        const fields = mapScrapedItemToSharePointFields(listName, data[i], sharepointConfig);

        // Title is required for most lists; enforce minimal safety
        if (!fields.Title) {
          fields.Title = data[i]?.title || `Item ${i + 1}`;
        }

        await createListItem(accessToken, siteId, listId, fields);

        successCount++;
        if ((i + 1) % 10 === 0) {
          console.log(`   Progress: ${i + 1}/${data.length} items inserted`);
        }
      } catch (itemErr) {
        errorCount++;
        console.error(`   Error inserting item ${i + 1}:\n${itemErr.message}`);
      }
    }

    console.log(`✅ Inserted ${successCount} items into ${listName} (${errorCount} errors)`);
  } catch (err) {
    console.error(`❌ Error inserting into ${listName}:\n${err.message}`);
  }
}

/**
 * Reset caches (should be called at the start of each run)
 */
function resetCaches() {
  cachedSiteId = null;
  cachedListIds.clear();
}

module.exports = {
  insertIntoSharePointList,
  resetCaches,
};
