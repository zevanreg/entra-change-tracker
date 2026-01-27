"""
SharePoint integration using Microsoft Graph API
Handles inserting scraped data into SharePoint lists
"""

import json
from typing import Dict, Any, List, Optional
from urllib.parse import quote

import requests

# Graph caches
_cached_site_id: Optional[str] = None
_cached_list_ids: Dict[str, str] = {}


def graph_fetch(token: str, method: str, graph_url: str, body: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    Make a Microsoft Graph API call.
    
    Args:
        token: Access token
        method: HTTP method
        graph_url: Graph API URL
        body: Request body
        
    Returns:
        Response JSON
    """
    headers = {
        'Authorization': f'Bearer {token}',
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    }
    
    response = requests.request(
        method=method,
        url=graph_url,
        headers=headers,
        json=body
    )
    
    if not response.ok:
        text = response.text
        raise RuntimeError(
            f"Graph HTTP {response.status_code} {response.reason}\n"
            f"URL: {graph_url}\n{text}"
        )
    
    return response.json()


def get_site_id_from_site_url(token: str, site_url: str) -> str:
    """
    Get SharePoint site ID from site URL.
    
    Args:
        token: Access token
        site_url: SharePoint site URL
        
    Returns:
        Site ID
    """
    global _cached_site_id
    
    if _cached_site_id:
        return _cached_site_id
    
    from urllib.parse import urlparse
    
    parsed = urlparse(site_url)
    hostname = parsed.hostname  # e.g. m365j556631.sharepoint.com
    site_path = parsed.path.rstrip('/')  # e.g. /sites/EntraChangeTrackers
    
    # GET /sites/{hostname}:{server-relative-path}
    endpoint = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
    site = graph_fetch(token, 'GET', endpoint)
    
    if not site.get('id'):
        raise RuntimeError(f"Could not resolve siteId for {site_url}")
    
    _cached_site_id = site['id']
    return _cached_site_id


def get_list_id_by_title(token: str, site_id: str, list_title: str) -> str:
    """
    Get SharePoint list ID by title.
    
    Args:
        token: Access token
        site_id: Site ID
        list_title: List title
        
    Returns:
        List ID
    """
    global _cached_list_ids
    
    if list_title in _cached_list_ids:
        return _cached_list_ids[list_title]
    
    safe_title = list_title.replace("'", "''")
    endpoint = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
        f"?$filter=displayName eq '{safe_title}'&$select=id,displayName"
    )
    
    result = graph_fetch(token, 'GET', endpoint)
    matches = [x for x in result.get('value', []) if x.get('displayName') == list_title]
    
    if not matches:
        raise RuntimeError(f'List not found by title "{list_title}" (check list name/spelling).')
    
    list_id = matches[0]['id']
    _cached_list_ids[list_title] = list_id
    return list_id


def create_list_item(token: str, site_id: str, list_id: str, fields: Dict[str, Any]) -> Dict[str, Any]:
    """
    Create a list item in SharePoint.
    
    Args:
        token: Access token
        site_id: Site ID
        list_id: List ID
        fields: Item fields
        
    Returns:
        Created item
    """
    endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"
    return graph_fetch(token, 'POST', endpoint, {'fields': fields})


def item_exists(
    token: str,
    site_id: str,
    list_id: str,
    title: str,
    date_field: str,
    date_value: str
) -> bool:
    """
    Check if an item with the same title and date already exists in the list.
    
    Args:
        token: Access token
        site_id: Site ID
        list_id: List ID
        title: Item title
        date_field: Name of the date field (e.g., 'ReleaseDate', 'AnnouncementDate')
        date_value: Date value to check
        
    Returns:
        True if item exists, false otherwise
    """
    if not title or not date_value:
        return False
    
    try:
        # Escape single quotes in title for OData filter
        safe_title = title.replace("'", "''")
        safe_date = date_value.replace("'", "''")
        
        # Query for items with matching Title and date field
        endpoint = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"
            f"?$filter=fields/Title eq '{safe_title}' and fields/{date_field} eq '{safe_date}'"
            f"&$select=id&$top=1"
        )
        
        result = graph_fetch(token, 'GET', endpoint)
        return result.get('value') and len(result['value']) > 0
    except Exception as err:
        # If query fails, assume item doesn't exist to avoid blocking insertion
        print(f"   Could not check for duplicate: {err}")
        return False


def map_scraped_item_to_sharepoint_fields(
    list_name: str,
    item: Dict[str, Any],
    config: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Map scraped item to SharePoint fields.
    
    Args:
        list_name: List name (e.g., 'EntraRoadmapItems')
        item: Scraped item
        config: SharePoint configuration with sharePointFieldMappings
        
    Returns:
        Mapped fields
    """
    # Collect lists from both browserScraping and httpScraping sections
    list_entries = {}
    if config:
        browser_scraping = config['browserScraping']
        if 'roadmap' in browser_scraping:
            list_entries['roadmap'] = browser_scraping['roadmap']
        if 'changeAnnouncements' in browser_scraping:
            list_entries['changeAnnouncements'] = browser_scraping['changeAnnouncements']
        
        http_scraping = config['httpScraping']
        sharepoint_lists = http_scraping['sharepointList']
        if 'whatsNew' in sharepoint_lists:
            list_entries['whatsNew'] = sharepoint_lists['whatsNew']
    
    list_key = None
    for key, entry in list_entries.items():
        sharepoint_list = entry['sharepointList'] if key != 'whatsNew' else entry
        name = sharepoint_list['name']
        if name.lower() == list_name.lower():
            list_key = key
            break

    # Get mapping from list config
    mapping = None
    if list_key and list_key in list_entries:
        entry = list_entries[list_key]
        sharepoint_list = entry['sharepointList'] if key != 'whatsNew' else entry
        mapping = sharepoint_list['mapping']
    
    if not mapping:
        raise RuntimeError(
            f'No SharePoint field mapping found for list "{list_name}". '
            f'Please define lists.<key>.mapping (or legacy sharePointFieldMappings) in config.json.'
        )
    
    # Map fields according to config
    fields = {}
    for sp_internal_name, source_key in mapping.items():
        fields[sp_internal_name] = item.get(source_key, '')
    
    # Ensure Title exists if possible
    if 'Title' not in fields and 'title' in item:
        fields['Title'] = item['title']
    
    return fields


def insert_into_sharepoint_list(
    list_name: str,
    data: List[Dict[str, Any]],
    access_token: str,
    sharepoint_config: Dict[str, Any]
) -> None:
    """
    Insert data into a SharePoint list using Microsoft Graph API.
    
    Args:
        list_name: The name of the SharePoint list
        data: The array of data items to insert
        access_token: Access token for Graph API
        sharepoint_config: SharePoint configuration
    """
    if not sharepoint_config or not access_token:
        print(f"⏭️ Skipping SharePoint insertion for {list_name} (not configured)")
        return
    
    try:
        print(f"📤 Inserting {len(data)} items into SharePoint list (Graph): {list_name}")
        
        site_id = get_site_id_from_site_url(access_token, sharepoint_config['sharepoint']['siteUrl'])
        list_id = get_list_id_by_title(access_token, site_id, list_name)
        
        # Determine the date field name based on list config
        list_name_lower = list_name.lower()
        date_field = None
        
        # Collect lists from browserScraping and httpScraping
        list_entries = {}
        browser_scraping = sharepoint_config['browserScraping']
        if 'roadmap' in browser_scraping:
            list_entries['roadmap'] = browser_scraping['roadmap']
        if 'changeAnnouncements' in browser_scraping:
            list_entries['changeAnnouncements'] = browser_scraping['changeAnnouncements']
        
        http_scraping = sharepoint_config['httpScraping']
        sharepoint_lists = http_scraping['sharepointList']
        if 'whatsNew' in sharepoint_lists:
            list_entries['whatsNew'] = sharepoint_lists['whatsNew']
        
        for key, entry in list_entries.items():
            sharepoint_list = entry['sharepointList'] if key != 'whatsNew' else entry
            name = sharepoint_list['name']
            if name.lower() == list_name_lower:
                date_field = sharepoint_list['dateField']
                break
        
        success_count = 0
        error_count = 0
        skipped_count = 0
        
        for i, item in enumerate(data):
            try:
                fields = map_scraped_item_to_sharepoint_fields(list_name, item, sharepoint_config)
                
                # Title is required for most lists; enforce minimal safety
                if 'Title' not in fields:
                    fields['Title'] = item.get('title', f'Item {i + 1}')
                
                # Check if item already exists (if date field is available)
                if date_field and date_field in fields:
                    exists = item_exists(
                        access_token,
                        site_id,
                        list_id,
                        fields['Title'],
                        date_field,
                        fields[date_field]
                    )
                    
                    if exists:
                        skipped_count += 1
                        if (i + 1) % 10 == 0:
                            print(f"   Progress: {i + 1}/{len(data)} processed ({skipped_count} duplicates skipped)")
                        continue
                
                create_list_item(access_token, site_id, list_id, fields)
                
                success_count += 1
                if (i + 1) % 10 == 0:
                    print(f"   Progress: {i + 1}/{len(data)} processed ({success_count} inserted, {skipped_count} duplicates)")
            except Exception as item_err:
                error_count += 1
                print(f"   Error inserting item {i + 1}:\n{item_err}")
        
        print(f"✅ Inserted {success_count} items into {list_name} ({skipped_count} duplicates skipped, {error_count} errors)")
    except Exception as err:
        print(f"❌ Error inserting into {list_name}:\n{err}")


def reset_caches() -> None:
    """Reset caches (should be called at the start of each run)."""
    global _cached_site_id, _cached_list_ids
    _cached_site_id = None
    _cached_list_ids = {}
