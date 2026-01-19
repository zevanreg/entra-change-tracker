# Entra Playwright Scraper

Automated web scraping tool for extracting Microsoft Entra (Azure AD) roadmap items and change announcements using Playwright. Optionally syncs data to SharePoint lists.

## Features

- 🔍 Scrapes Entra Change Management Hub (Roadmap & Change Announcements)
- 📊 Extracts detailed information including descriptions and URLs
- 💾 Saves data locally as timestamped JSON files
- 📤 Optional SharePoint integration using Microsoft Graph API
- 🎯 Date range filtering support
- 🔐 Persistent browser context authentication

## Prerequisites

- **Node.js** (v16 or higher)
- **Microsoft Edge** browser installed
- **Azure AD account** with access to Entra portal
- **(Optional)** SharePoint site with appropriate permissions

## Installation

### 1. Clone or download this repository

```bash
cd c:\repos\entra-change-tracker\js
   {
     "siteUrl": "https://yourtenant.sharepoint.com/sites/yoursite",
     "clientId": "your-app-client-id",
     "tenantId": "your-tenant-id",
     "dateFilter": "Last 3 months",
     "saveToFile": false,
     "lists": {
       "roadmap": {
         "name": "EntraRoadmapItems",
         "dateField": "ReleaseDate",
         "mapping": {
           "Title": "title",
           "Category": "changeEntityCategory",
           "Service": "changeEntityService",
           "ReleaseType": "changeEntityDeliveryStage",
           "ReleaseDate": "publishStartDateTime",
           "State": "changeEntityState",
           "Overview": "overview",
           "Description": "description",
           "Url": "url"
         }
       },
       "changeAnnouncements": {
         "name": "EntraChangeAnnouncements",
         "dateField": "AnnouncementDate",
         "mapping": {
           "Title": "title",
           "Service": "changeEntityService",
           "ChangeType": "changeEntityChangeType",
           "AnnouncementDate": "announcementDateTime",
           "TargetDate": "targetDateTime",
           "ActionRequired": "isCustomerActionRequired",
           "Tags": "marketingThemes",
           "Overview": "overview",
           "Description": "description",
           "Url": "url"
         }
       }
     }
   }
- `@azure/msal-node-extensions` - Token cache support

### 3. Install Playwright browsers (if needed)

```bash
npx playwright install msedge
```

## Configuration

### SharePoint Configuration (Optional)

If you want to sync data to SharePoint:

1. **Copy the template configuration:**
  ```bash
  copy config.json.template config.json
  ```

2. **Create an Azure AD App Registration:**
   - Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
   - Create a new registration
   - Note the **Application (client) ID**
  - Grant **Microsoft Graph API permissions**: `Sites.ReadWrite.All`

3. **Edit `config.json`:**
   ```json
   {
     "siteUrl": "https://yourtenant.sharepoint.com/sites/yoursite",
     "clientId": "your-app-client-id",
     "tenantId": "your-tenant-id",
     "dateFilter": "Last 3 months",
     "saveToFile": false,
     "lists": {
       "roadmap": "Roadmap",
       "changeAnnouncements": "ChangeAnnouncements"
     }
   }
   ```

4. **Create SharePoint lists** with these columns:
   - **Roadmap**: Title, Category, Service, ReleaseType, ReleaseDate, State, URL, Description
   - **ChangeAnnouncements**: Title, Category, Service, ReleaseType, ReleaseDate, State, URL, Description

### Date Filter Options

Valid values for `dateFilter` in config:
- `"Last 1 month"`
- `"Last 3 months"`
- `"Last 6 months"`
- `"Last 1 year"`
- `""` (empty for all results visible by default)

### File Output

Control local JSON output with `saveToFile` in config:
- `true` (default) saves timestamped JSON files
- `false` skips saving local files

## Usage

Run the scraper using Edge's persistent browser context:

```bash
node entra.js
```

**First run:**
1. Browser will open (non-headless)
2. Log in to your Microsoft account
3. Script will scrape data after login
4. Credentials are saved in `edge-profile/` directory

**Subsequent runs:**
- Authentication is maintained
- Set `headless: true` in the script if desired

## Output

### Local Files

Data is saved to timestamped JSON files:
- `roadmap-2026-01-14T10-30-00.json`
- `change-announcements-2026-01-14T10-30-00.json`

### SharePoint Lists

If configured, data is automatically inserted into specified SharePoint lists with progress tracking.

## Project Structure

```
entra-playwright/
├── entra.js                         # Main scraper
├── config.json                      # SharePoint & filter config (gitignored)
├── config.json.template             # Configuration template
├── package.json                     # Node.js dependencies
├── .gitignore                       # Git ignore rules
├── edge-profile/                    # Edge browser profile (gitignored)
└── README.md                        # This file
```

## How It Works

1. **Authentication**: Uses persistent browser context to maintain login session
2. **Navigation**: Opens Entra Change Management Hub
3. **Tab Switching**: Clicks Roadmap and Change Announcements tabs
4. **Date Filtering**: Applies date range filter from config
5. **Scrolling**: Handles virtualized lists by scrolling and capturing rows
6. **Detail Extraction**: Clicks each row to extract full details
7. **Data Export**: Saves to JSON and optionally to SharePoint

## Troubleshooting

### Authentication Issues

**Problem**: Login prompt appears or authentication fails

**Solutions:**
- Delete `edge-profile/` directory and run script again to re-authenticate
- Ensure you have access to Entra portal
- Verify your Microsoft account credentials are valid

### SharePoint Connection Issues

**Problem**: "Error inserting into SharePoint list"

**Solutions:**
- Verify `config.json` credentials are correct
- Check Azure AD app has `Sites.ReadWrite.All` permission
- Ensure SharePoint lists exist with correct names
- Verify list columns match the data structure

### Scraping Incomplete Data

**Problem**: Not all items are captured

**Solutions:**
- Increase `TOTAL_TIMEOUT_MS` in scraping function
- Adjust `PASS_DELAY_MS` for slower page rendering
- Check network connection stability
- Verify date filter is applied correctly

### Rate Limiting

**Problem**: SharePoint insertion fails with 429 errors

**Solutions:**
- Add delays between SharePoint insertions
- Reduce batch size
- Use SharePoint batch operations (future enhancement)

## Development

### Running in Debug Mode

Set `headless: false` in the script to see browser actions:

```javascript
const context = await chromium.launchPersistentContext('./edge-profile', {
  channel: 'msedge',
  headless: false, // Watch the automation
  // ...
});
```

### Modifying Selectors

If Entra UI changes, update selectors in:
- `clickTab()` function
- `scrapeDetailsList()` function
- `extractRowDetails()` function

## Security Notes

⚠️ **Important Security Considerations:**

- Never commit `config.json`
- Never commit `edge-profile/` directory (contains credentials)
- These are already excluded in `.gitignore`
- Rotate client secrets regularly
- Use least-privilege permissions for Azure AD app

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## Support

For issues or questions:
1. Check the Troubleshooting section
2. Review Playwright documentation: https://playwright.dev
3. Review PnPjs documentation: https://pnp.github.io/pnpjs/
