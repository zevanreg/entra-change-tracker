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

2. **Create an Azure AD App Registration** (for device code flow):
   - Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
   - Create a new registration
   - Note the **Application (client) ID** and **Tenant ID**
   - Enable **Allow public client flows** (Authentication → Advanced settings)
   - Grant **Microsoft Graph API permissions**: `Sites.ReadWrite.All`
   - Admin consent may be required

3. **Edit `config.json`** - See Configuration Reference below for all settings
4. **Create SharePoint lists** with required columns (see Configuration Reference below)

## Configuration Reference

The `config.json` file uses a hierarchical structure organized into three main sections:

### Structure Overview

```json
{
  "sharepoint": { /* SharePoint and authentication settings */ },
  "browserScraping": { /* Entra Portal scraping configuration */ },
  "httpScraping": { /* What's New public page scraping */ }
}
```

### SharePoint Section

Configures SharePoint site connection and authentication:

```json
"sharepoint": {
  "siteUrl": "https://yourtenant.sharepoint.com/sites/yoursite",
  "authentication": {
    "authMethod": "devicecode",
    "devicecode": {
      "clientId": "your-app-client-id",
      "tenantId": "your-tenant-id"
    }
  }
}
```

**Settings:**
- `siteUrl` (string, required): Full URL to your SharePoint site
- `authentication.authMethod` (string): Authentication method - `"devicecode"` or `"iwa"` (Integrated Windows Authentication)
- `authentication.devicecode.clientId` (string): Azure AD app client ID (required for device code flow)
- `authentication.devicecode.tenantId` (string): Azure AD tenant ID (required for device code flow)

**Authentication Methods:**
- **Device Code Flow** (`"devicecode"`): Interactive browser-based authentication, requires app registration
- **Integrated Windows Authentication** (`"iwa"`): Uses current Windows credentials, no app registration needed

### Browser Scraping Section

Configures Playwright-based scraping of the Entra Portal:

```json
"browserScraping": {
  "entraPortal": "https://entra.microsoft.com/",
  "dateFilter": "Last 3 months",
  "selectors": { /* CSS selectors */ },
  "timeouts": { /* Timeout values */ },
  "scraperConfig": { /* Scraping behavior */ },
  "textPatterns": { /* Text matching patterns */ },
  "roadmap": {
    "saveToFile": false,
    "extractDetails": true,
    "sharepointList": { /* List configuration */ }
  },
  "changeAnnouncements": { /* Same structure as roadmap */ }
}
```

**General Settings:**
- `entraPortal` (string): Entra portal URL
- `dateFilter` (string): Date range filter - `"Last 1 month"`, `"Last 3 months"`, `"Last 6 months"`, `"Last 1 year"`, or `""` (all)
- `selectors` (object): CSS selectors for page elements (see template for details)
- `timeouts` (object): Timeout values in milliseconds for various operations
- `scraperConfig` (object): Scraping behavior settings (delays, retries, etc.)
- `textPatterns` (object): Regular expressions for text matching

**Per-Source Settings (roadmap, changeAnnouncements):**
- `saveToFile` (boolean): Whether to save data as timestamped JSON file
- `extractDetails` (boolean): Whether to click into each item for full details (slower but more complete)
- `sharepointList.name` (string): SharePoint list name
- `sharepointList.dateField` (string): Field name for date column
- `sharepointList.mapping` (object): Maps SharePoint columns to data fields

**Required SharePoint List Columns:**
- **Roadmap**: Title, Category, Service, ReleaseType, ReleaseDate, State, Overview, Description, Url
- **Change Announcements**: Title, Service, ChangeType, AnnouncementDate, TargetDate, ActionRequired, Tags, Overview, Description, Url

### HTTP Scraping Section

Configures HTTP-based scraping of the public What's New page:

```json
"httpScraping": {
  "whatsNew": "https://learn.microsoft.com/en-us/entra/fundamentals/whats-new",
  "microsoftLearnBase": "https://learn.microsoft.com",
  "saveToFile": false,
  "releaseTypeMapping": { /* Keyword to release type mapping */ },
  "sharepointList": {
    "whatsNew": {
      "name": "EntraWhatsNew",
      "dateField": "PublishDate",
      "mapping": { /* Column mappings */ }
    }
  }
}
```

**Settings:**
- `whatsNew` (string): URL to What's New page
- `microsoftLearnBase` (string): Base URL for resolving relative links
- `saveToFile` (boolean): Whether to save data as timestamped JSON file
- `releaseTypeMapping` (object): Maps keywords in titles to release types (e.g., "preview" → "Public Preview")
- `sharepointList.whatsNew.name` (string): SharePoint list name
- `sharepointList.whatsNew.dateField` (string): Field name for date column
- `sharepointList.whatsNew.mapping` (object): Maps SharePoint columns to data fields

**Required SharePoint List Columns:**
- **What's New**: Title, PublishDate, Category, Url, Description, ReleaseType

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
js/
├── entra.js                         # Main entry point
├── config.json                      # Configuration (create from template)
├── config.json.template             # Configuration template
├── package.json                     # Node.js dependencies
├── .gitignore                       # Git ignore rules
├── pw-profile/                      # Playwright browser profile (auto-generated)
├── README.md                        # This file
└── lib/
    ├── config.js                    # Configuration management
    ├── auth.js                      # Authentication (device code/IWA)
    ├── browser-helpers.js           # Browser automation helpers
    ├── scraper.js                   # Main scraping logic
    └── sharepoint.js                # SharePoint/Graph API integration
```

## How It Works

### Architecture

The scraper uses a modular architecture with separation of concerns:

- **config.js**: Loads and validates configuration once at startup
- **auth.js**: Handles authentication (device code flow or IWA) and token management
- **browser-helpers.js**: Provides lazy-loaded browser automation utilities
- **scraper.js**: Coordinates scraping operations for Entra portal and What's New pages
- **sharepoint.js**: Manages Microsoft Graph API interactions for SharePoint
- **entra.js**: Main orchestrator that coordinates all modules

### Execution Flow

1. **Initialization**: Loads configuration and authenticates to Microsoft Graph
2. **Browser Context**: Uses persistent browser profile to maintain Entra portal session
3. **Navigation**: Opens Entra Change Management Hub
4. **Tab Switching**: Clicks Roadmap and Change Announcements tabs
5. **Date Filtering**: Applies configured date range filter
6. **Scrolling**: Handles virtualized lists by scrolling and capturing rows
7. **Detail Extraction**: Clicks each row to extract full details
8. **HTTP Scraping**: Fetches What's New items from public Microsoft Learn page
9. **Data Export**: Saves to JSON files and optionally inserts into SharePoint lists

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
