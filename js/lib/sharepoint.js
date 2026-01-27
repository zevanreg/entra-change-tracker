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
 * Check if an item with the same title and date already exists in the list
 * @param {string} token - Access token
 * @param {string} siteId - Site ID
 * @param {string} listId - List ID
 * @param {string} title - Item title
 * @param {string} dateField - Name of the date field (e.g., 'ReleaseDate', 'AnnouncementDate')
 * @param {string} dateValue - Date value to check
 * @returns {Promise<boolean>} True if item exists, false otherwise
 */
async function itemExists(token, siteId, listId, title, dateField, dateValue) {
  if (!title || !dateValue) return false;
  
  try {
    // Escape single quotes in title for OData filter
    const safeTitle = title.replace(/'/g, "''");
    const safeDate = dateValue.replace(/'/g, "''");
    
    // Query for items with matching Title and date field
    const endpoint = 
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items` +
      `?$filter=fields/Title eq '${safeTitle}' and fields/${dateField} eq '${safeDate}'` +
      `&$select=id&$top=1`;
    
    const result = await graphFetch(token, "GET", endpoint);
    return result.value && result.value.length > 0;
  } catch (err) {
    // If query fails, assume item doesn't exist to avoid blocking insertion
    console.warn(`   Could not check for duplicate: ${err.message}`);
    return false;
  }
}

/**
 * Map scraped item to SharePoint fields
 * @param {string} listName - List name (e.g., 'EntraRoadmapItems')
 * @param {object} item - Scraped item
 * @param {object} config - Configuration
 * @returns {object} Mapped fields
 */
function mapScrapedItemToSharePointFields(listName, item, config) {
  // Collect lists from both browserScraping and httpScraping sections
  const listEntries = {};
  if (config) {
    const browserScraping = config.browserScraping;
    if (browserScraping.roadmap) {
      listEntries.roadmap = browserScraping.roadmap;
    }
    if (browserScraping.changeAnnouncements) {
      listEntries.changeAnnouncements = browserScraping.changeAnnouncements;
    }
    
    const httpScraping = config.httpScraping;
    const sharepointLists = httpScraping.sharepointList;
    if (sharepointLists.whatsNew) {
      listEntries.whatsNew = sharepointLists.whatsNew;
    }
  }
  
  const listKey = Object.keys(listEntries).find((key) => {
    const entry = listEntries[key];
    const sharepointList = key === 'whatsNew' ? entry : entry.sharepointList;
    return sharepointList.name.toLowerCase() === listName.toLowerCase();
  });

  // Get mapping from list config
  const entry = listEntries[listKey];
  const sharepointList = listKey === 'whatsNew' ? entry : entry.sharepointList;
  const mapping = sharepointList.mapping;
  
  if (!mapping) {
    throw new Error(
      `No SharePoint field mapping found for list "${listName}". ` +
      `Please define lists.<key>.mapping (or legacy sharePointFieldMappings) in config.json.`
    );
  }
  
  // Map fields according to config
  const fields = {};
  for (const [spInternalName, sourceKey] of Object.entries(mapping)) {
    fields[spInternalName] = item[sourceKey] || "";
  }
  
  // Ensure Title exists if possible
  if (!fields.Title && item.title) {
    fields.Title = item.title;
  }
  
  return fields;
}

/**
 * Insert data into a SharePoint list using Microsoft Graph API
 * @param {string} listName - The name of the SharePoint list
 * @param {object[]} data - The array of data items to insert
 * @param {string} accessToken - Access token for Graph API
 * @param {object} config - Configuration
 * @returns {Promise<void>}
 */
async function insertIntoSharePointList(listName, data, accessToken, config) {
  if (!config || !accessToken) {
    console.log(`⏭️ Skipping SharePoint insertion for ${listName} (not configured)`);
    return;
  }

  try {
    console.log(`📤 Inserting ${data.length} items into SharePoint list (Graph): ${listName}`);

    const siteId = await getSiteIdFromSiteUrl(accessToken, config.sharepoint.siteUrl);
    const listId = await getListIdByTitle(accessToken, siteId, listName);

    // Determine the date field name based on list config
    const listNameLower = listName.toLowerCase();
    
    // Collect lists from browserScraping and httpScraping
    const listEntries = {};
    const browserScraping = config.browserScraping;
    if (browserScraping.roadmap) {
      listEntries.roadmap = browserScraping.roadmap;
    }
    if (browserScraping.changeAnnouncements) {
      listEntries.changeAnnouncements = browserScraping.changeAnnouncements;
    }
    
    const httpScraping = config.httpScraping;
    const sharepointLists = httpScraping.sharepointList;
    if (sharepointLists.whatsNew) {
      listEntries.whatsNew = sharepointLists.whatsNew;
    }
    
    const listKey = Object.keys(listEntries).find((key) => {
      const entry = listEntries[key];
      const sharepointList = key === 'whatsNew' ? entry : entry.sharepointList;
      return sharepointList.name.toLowerCase() === listNameLower;
    });
    
    let dateField = null;
    if (listKey) {
      const entry = listEntries[listKey];
      const sharepointList = listKey === 'whatsNew' ? entry : entry.sharepointList;
      dateField = sharepointList.dateField;
    }

    let successCount = 0;
    let errorCount = 0;
    let skippedCount = 0;

    for (let i = 0; i < data.length; i++) {
      try {
        const fields = mapScrapedItemToSharePointFields(listName, data[i], config);

        // Title is required for most lists; enforce minimal safety
        if (!fields.Title) {
          fields.Title = data[i].title || `Item ${i + 1}`;
        }

        // Check if item already exists (if date field is available)
        if (dateField && fields[dateField]) {
          const exists = await itemExists(
            accessToken,
            siteId,
            listId,
            fields.Title,
            dateField,
            fields[dateField]
          );
          
          if (exists) {
            skippedCount++;
            if ((i + 1) % 10 === 0) {
              console.log(`   Progress: ${i + 1}/${data.length} processed (${skippedCount} duplicates skipped)`);
            }
            continue;
          }
        }

        await createListItem(accessToken, siteId, listId, fields);

        successCount++;
        if ((i + 1) % 10 === 0) {
          console.log(`   Progress: ${i + 1}/${data.length} processed (${successCount} inserted, ${skippedCount} duplicates)`);
        }
      } catch (itemErr) {
        errorCount++;
        console.error(`   Error inserting item ${i + 1}:\n${itemErr.message}`);
      }
    }

    console.log(`✅ Inserted ${successCount} items into ${listName} (${skippedCount} duplicates skipped, ${errorCount} errors)`);
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
