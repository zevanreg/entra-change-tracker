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
```

### 2. Create a virtual environment (recommended)

```bash
python -m venv venv

# On Windows
venv\Scripts\activate

# On macOS/Linux
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

2. **Create an Azure AD App Registration:**
   - Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
   - Create a new registration
   - Note the **Application (client) ID** and **Tenant ID**
   - Grant **Microsoft Graph API permissions**: `Sites.ReadWrite.All`
   - Admin consent may be required

3. **Edit `config.json`:**
   ```json
   {
     "siteUrl": "https://yourtenant.sharepoint.com/sites/yoursite",
     "clientId": "your-app-client-id",
     "tenantId": "your-tenant-id",
     "dateFilter": "Last 3 months",
     "lists": {
       "roadmap": "EntraRoadmapItems",
       "changeAnnouncements": "EntraChangeAnnouncements"
     }
   }
   ```

4. **Create SharePoint lists** with these columns:
   - **EntraRoadmapItems**: Title, Category, Service, ReleaseType, ReleaseDate, State, URL, Description, Overview
   - **EntraChangeAnnouncements**: Title, Service, ChangeType, AnnouncementDate, TargetDate, ActionRequired, Tags, URL, Description, Overview

### Date Filter Options

Valid values for `dateFilter` in config:
- `"Last 1 month"`
- `"Last 3 months"`
- `"Last 6 months"`
- `"Last 1 year"`
- `""` (empty for all results)

## Usage

Run the scraper using Edge's persistent browser context:

```bash
python entra.py
```

**First run:**
1. Browser will open (non-headless)
2. Log in manually to your Microsoft account
3. Browser profile will be saved in `./python-edge-profile/`
4. If SharePoint is configured, you'll see a device code authentication prompt

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
- Delete `.token-cache.json` to force re-authentication
- Verify clientId and tenantId in `config.json`
- Ensure app has proper permissions in Azure AD

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
- Uses `msal` library instead of `@azure/msal-node`
- Uses `playwright` async API
- Uses `requests` library for HTTP calls

## License

This is a port of the original JavaScript version with identical functionality.
