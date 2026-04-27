# Entra Change Tracker

Automated tool for tracking Microsoft Entra (Azure AD) changes — scraping roadmap items, change announcements, and What's New entries from the Entra portal and Microsoft Learn. Optionally syncs collected data to SharePoint lists via the Microsoft Graph API.

## Features

- Scrapes the Entra Change Management Hub (Roadmap & Change Announcements)
- Fetches What's New items from the public Microsoft Learn page
- Saves results locally as timestamped JSON files
- Optional SharePoint integration for centralised tracking
- Date-range filtering and persistent browser authentication

## Implementations

Choose the version that best fits your environment:

| Version | Folder | Requirements |
|---------|--------|--------------|
| **JavaScript (Node.js)** | [`js/`](./js/) | Node.js 16+, Microsoft Edge |
| **Python** | [`python/`](./python/) | Python 3.8+, Microsoft Edge |

Both versions are functionally equivalent. See the README in each folder for installation, configuration, and usage instructions.
