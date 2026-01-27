# Entra Playwright Scraper (Python)

Automated web scraping tool for extracting Microsoft Entra (Azure AD) roadmap items and change announcements using Playwright. Optionally syncs data to SharePoint lists.

## Features

- 🔍 Scrapes Entra Change Management Hub (Roadmap & Change Announcements)
- 📊 Extracts detailed information including descriptions and URLs
- 💾 Saves data locally as timestamped JSON files
- 📤 Optional SharePoint integration using Microsoft Graph API
- 🎯 Date range filtering support
- 🔐 Persistent browser context authentication

## Prerequisites

- **Python** (3.8 or higher)
- **Microsoft Edge** browser installed
- **Azure AD account** with access to Entra portal
- **(Optional)** SharePoint site with appropriate permissions

## Installation

### 1. Navigate to the Python folder

```bash
cd c:\repos\entra-change-tracker\python
source venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Install Playwright browsers

```bash
playwright install msedge
```

## Configuration

### SharePoint Configuration (Optional)

If you want to sync data to SharePoint:

1. **Copy the template configuration:**
   ```bash
   copy config.json.template config.json
   ```

2. **Choose authentication method:**
   
   **Option A: Device Code Flow**
   - Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
   - Create a new registration
   - Note the **Application (client) ID** and **Tenant ID**
   - Enable **Allow public client flows** (Authentication → Advanced settings)
   - Grant **Microsoft Graph API permissions**: `Sites.ReadWrite.All`
   - Admin consent may be required

   **Option B: Integrated Windows Authentication (IWA)**
   - Uses your current Windows credentials automatically
   - No app registration required
   - Requires appropriate Azure AD permissions on your account
   - Ideal for corporate environments with SSO

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
python entra.py
```

**First run:**
1. Browser will open (non-headless)
2. Log in manually to your Microsoft account
3. Browser profile will be saved in `./edge-profile/`
4. If SharePoint is configured:
   - **Device code flow**: You'll see a device code authentication prompt
   - **IWA**: Authentication happens automatically using your Windows credentials

**Subsequent runs:**
- Browser will use saved profile (no login required)
- Access token will be cached in `.token-cache.json`

## Output

### JSON Files

Data is saved in the `python` folder with timestamps:
- `roadmap-YYYY-MM-DDTHH-MM-SS.json`
- `change-announcements-YYYY-MM-DDTHH-MM-SS.json`

### SharePoint Lists (if configured)

Items are automatically inserted into configured SharePoint lists with duplicate checking.

## Project Structure

```
python/
├── entra.py                    # Main entry point
├── config.json                 # Configuration (create from template)
├── config.json.template        # Configuration template
├── requirements.txt            # Python dependencies
├── .token-cache.json          # MSAL token cache (auto-generated)
├── python-edge-profile/       # Edge browser profile (auto-generated)
└── lib/
    ├── __init__.py
    ├── auth.py                # Authentication and config loading
    ├── browser_helpers.py     # Browser automation helpers
    ├── scraper.py             # Main scraping logic
    └── sharepoint.py          # SharePoint/Graph API integration
```

## Troubleshooting

### Browser doesn't open
- Ensure Microsoft Edge is installed
- Run `playwright install msedge`

### Authentication issues

**Device Code Flow:**
- Delete `.token-cache.json` to force re-authentication
- Verify `clientId` and `tenantId` in `config.json`
- Ensure app has proper permissions in Azure AD
- Ensure "Allow public client flows" is enabled in app registration

**Integrated Windows Authentication (IWA):**
- Ensure you're logged into Windows with your Azure AD account
- May require Azure CLI installed: `pip install azure-cli` or download from Microsoft
- Verify `tenantId` in `config.json`
- Check that your account has appropriate SharePoint permissions
- Try setting `authMethod` to `"devicecode"` if IWA doesn't work in your environment

### SharePoint insertion fails
- Verify list names match exactly (case-sensitive)
- Check that required columns exist in SharePoint lists
- Ensure app has Sites.ReadWrite.All permissions

### "Failed to load config.json"
- Copy `config.json.template` to `config.json`
- Ensure JSON is valid (no trailing commas, proper quotes)

## Differences from JavaScript Version

This Python version maintains feature parity with the JavaScript version but with Python-specific implementations:
- Uses `asyncio` instead of native async/await
- Uses `msal` library for device code flow (instead of `@azure/msal-node`)
- Uses `azure-identity` library for IWA (instead of `@azure/identity`)
- Uses `playwright` async API
- Uses `requests` library for HTTP calls
- Supports both device code flow and IWA authentication methods

## License

This is a port of the original JavaScript version with identical functionality.
