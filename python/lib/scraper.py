"""
Entra Portal Scraper
Handles all browser automation and data extraction from Entra portal
"""

import re
from pathlib import Path
from typing import Dict, Any, List, Optional
from datetime import datetime

from playwright.async_api import async_playwright, Page, Frame
from bs4 import BeautifulSoup
import requests

from .browser_helpers import (
    wait_for_splash_screen,
    click_tab,
    set_date_range_filter,
    scrape_details_list
)
from .auth import get_configuration


def extract_release_type_from_title(title: str) -> tuple[str, str]:
    """
    Extract release type from the beginning of the title.
    
    Args:
        title: The full title string
        
    Returns:
        Tuple of (release_type, cleaned_title)
    """
    config = get_configuration().get('config', {})
    release_type_mapping = config.get('releaseTypeMapping', {})
    
    # Check if title starts with any known release type (case-insensitive)
    title_lower = title.lower()
    for page_value, mapped_value in release_type_mapping.items():
        if title_lower.startswith(page_value.lower()):
            # Extract the part after the release type
            remainder = title[len(page_value):].strip()
            # Remove common separators at the start
            for sep in ['-', ':', '–', '—']:
                if remainder.startswith(sep):
                    remainder = remainder[1:].strip()
                    break
            return mapped_value, remainder
    
    # No release type found - check if there's an unmapped one
    if any(sep in title[:50] for sep in [' - ', ': ', ' – ']):
        for sep in [' - ', ': ', ' – ']:
            if sep in title:
                potential_type = title.split(sep)[0].strip()
                if potential_type and potential_type[0].isupper():
                    print(f"⚠️ Unmapped release type found: '{potential_type}' in title: {title[:60]}...")
                break
    
    return '', title


async def scrape_tab(
    page: Page,
    frame: Frame,
    tab_name: str,
    date_filter: Optional[str] = None,
    extract_details: bool = True
) -> Optional[List[Dict[str, Any]]]:
    """
    Scrape data from a specific tab.
    
    Args:
        page: The page
        frame: The iframe
        tab_name: The name of the tab to scrape
        date_filter: Optional date filter to apply
        extract_details: Whether to extract details by clicking rows
        
    Returns:
        Scraped data or None if tab not found
    """
    # Click the tab
    tab_clicked = await click_tab(frame, tab_name)
    
    if not tab_clicked:
        print(f"❌ Could not locate the {tab_name} tab/menu.")
        return None
    
    # Set date filter if specified
    if date_filter:
        filter_set = await set_date_range_filter(frame, date_filter)
        if not filter_set:
            print('⚠️ Could not set date range filter, continuing anyway...')
    
    # Scrape the data
    data = await scrape_details_list(page, frame, extract_details)
    print(f"✅ Extracted {len(data)} items from {tab_name}")
    
    return data


def scrape_whats_new_page() -> Optional[List[Dict[str, Any]]]:
    """
    Scrape data from Microsoft Learn What's New page.
    
    Returns:
        List of items with Release Type, Title, Type, Service Category, 
        Product Capability, Detail, Link, Date
    """
    config = get_configuration().get('config', {})
    url = config.get('urls', {}).get('whatsNew', "https://learn.microsoft.com/en-us/entra/fundamentals/whats-new")
    
    try:
        print(f"🌐 Fetching {url}...")
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.content, 'html.parser')
        items = []
        
        # Find all h2 headers that represent month sections
        month_sections = soup.find_all('h2', id=True)
        
        for month_section in month_sections:
            month_text = month_section.get_text(strip=True)
            
            # Skip if not a date header (e.g., "January 2026")
            if not any(month in month_text for month in ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']):
                continue
            
            # Find all h3 items under this month
            current = month_section.find_next_sibling()
            while current and current.name != 'h2':
                if current.name == 'h3':
                    item = extract_whats_new_item(current, month_text)
                    if item:
                        items.append(item)
                current = current.find_next_sibling()
        
        print(f"✅ Extracted {len(items)} items from What's New page")
        return items
        
    except Exception as e:
        print(f"❌ Error scraping What's New page: {e}")
        return None


def extract_whats_new_item(h3_element, month_text: str) -> Optional[Dict[str, Any]]:
    """
    Extract a single item from the What's New page.
    
    Args:
        h3_element: The h3 element containing the item title
        month_text: The month/year text (e.g., "January 2026")
        
    Returns:
        Dictionary with item data
    """
    try:
        item = {
            'releaseType': '',
            'title': '',
            'type': '',
            'serviceCategory': '',
            'productCapability': '',
            'detail': '',
            'link': '',
            'date': ''
        }
        
        # Extract title and link
        title_link = h3_element.find('a')
        if title_link:
            full_title = title_link.get_text(strip=True)
            item['link'] = title_link.get('href', '')
            if item['link'] and not item['link'].startswith('http'):
                config = get_configuration().get('config', {})
                base_url = config.get('urls', {}).get('microsoftLearnBase', 'https://learn.microsoft.com')
                item['link'] = f"{base_url}{item['link']}"
        else:
            full_title = h3_element.get_text(strip=True)
        
        # Extract release type from title
        release_type, cleaned_title = extract_release_type_from_title(full_title)
        
        if not release_type and any(sep in full_title for sep in [' - ', ': ', ' – ']):
            # There might be an unmapped release type
            for sep in [' - ', ': ', ' – ']:
                if sep in full_title:
                    potential_type = full_title.split(sep)[0].strip()
                    # Check if it looks like a release type (title case, multiple words)
                    if potential_type and potential_type[0].isupper():
                        print(f"⚠️ Unmapped release type found: '{potential_type}' in title: {full_title[:60]}...")
                    break
        
        item['releaseType'] = release_type
        item['title'] = cleaned_title
        
        # Set date from month section
        item['date'] = month_text
        
        # Extract detail from the following paragraph(s)
        detail_parts = []
        current = h3_element.find_next_sibling()
        
        while current and current.name not in ['h2', 'h3']:
            if current.name == 'p':
                text = current.get_text(strip=True)
                if text:
                    # Look for metadata in strong tags
                    strong_tags = current.find_all('strong')
                    for strong in strong_tags:
                        label = strong.get_text(strip=True).rstrip(':')
                        # Get the text after the strong tag
                        next_text = strong.next_sibling
                        if next_text and isinstance(next_text, str):
                            value = next_text.strip().strip(':')
                            
                            if 'Type' in label and not item['type']:
                                item['type'] = value
                            elif 'Service category' in label or 'Service Category' in label:
                                item['serviceCategory'] = value
                            elif 'Product capability' in label or 'Product Capability' in label:
                                item['productCapability'] = value
                            elif 'Release' in label:
                                item['releaseType'] = value
                    
                    detail_parts.append(text)
            elif current.name == 'ul':
                # Add list items to detail
                list_items = current.find_all('li')
                for li in list_items:
                    detail_parts.append(f"• {li.get_text(strip=True)}")
            
            current = current.find_next_sibling()
        
        item['detail'] = ' '.join(detail_parts)
        
        # Only return item if it has a title
        return item if item['title'] else None
        
    except Exception as e:
        print(f"⚠️ Error extracting item: {e}")
        return None


async def scrape_entra_portal(
    date_filter: Optional[str] = None
) -> Dict[str, Optional[List[Dict[str, Any]]]]:
    """
    Scrape Roadmap data from Entra portal.
    
    Args:
        date_filter: Optional date filter
        
    Returns:
        Dictionary with roadmap and changeAnnouncements data
    """
    config = get_configuration().get('config', {})
    entra_url = config.get('urls', {}).get('entraPortal', "https://entra.microsoft.com/#blade/Microsoft_AAD_IAM/ChangeManagementHubList.ReactView")
    
    async with async_playwright() as playwright:
        profile_dir = Path(__file__).resolve().parent.parent / "edge-profile"
        edge_exe = Path("C:/Program Files/Microsoft/Edge/Application/msedge.exe")

        launch_kwargs = {
            "headless": False,
            "args": ["--disable-blink-features=AutomationControlled"],
            "viewport": {"width": 1920, "height": 1080},
            "device_scale_factor": 1,
        }

        if edge_exe.exists():
            launch_kwargs["executable_path"] = str(edge_exe)
        else:
            launch_kwargs["channel"] = "msedge"

        context = await playwright.chromium.launch_persistent_context(
            str(profile_dir),
            **launch_kwargs,
        )

        try:
            page = context.pages[0] if context.pages else await context.new_page()

            # Navigate to Entra portal
            await page.goto(entra_url, wait_until='domcontentloaded', timeout=60000)

            # Get the main iframe
            iframe_locator = page.locator('iframe[name="ChangeManagementHubList.ReactView"]')
            await iframe_locator.wait_for(state='attached', timeout=30000)

            iframe_handle = await iframe_locator.element_handle()
            frame = await iframe_handle.content_frame()

            if not frame:
                raise RuntimeError('ReactView frame attached, but content not available yet.')

            # Wait for initial splash screen to disappear
            await wait_for_splash_screen(frame, 60000)

            # Wait for progress dots to disappear
            try:
                progress_dots = frame.locator('div.fxs-progress-dots')
                await progress_dots.wait_for(state='hidden', timeout=15000)
            except Exception:
                # Progress dots might not exist or already hidden
                pass

            # Get config for extract_details settings
            config_obj = get_configuration().get('config', {})
            lists_config = config_obj.get('lists', {})

            # Scrape Roadmap
            roadmap_extract_details = lists_config.get('roadmap', {}).get('extractDetails', True)
            roadmap = await scrape_tab(page, frame, '/^Roadmap$/i', date_filter, roadmap_extract_details)

            # Scrape Change Announcements
            change_announcements_extract_details = lists_config.get('changeAnnouncements', {}).get('extractDetails', True)
            change_announcements = await scrape_tab(page, frame, '/^Change announcements$/i', date_filter, change_announcements_extract_details)

            return {
                'roadmap': roadmap,
                'changeAnnouncements': change_announcements
            }
        finally:
            await context.close()
            print('✅ Browser closed.')


async def scrape_all_sources(
    date_filter: Optional[str] = None
) -> Dict[str, Optional[List[Dict[str, Any]]]]:
    """
    Scrape data from all sources: Entra portal and Microsoft Learn What's New.
    
    Args:
        date_filter: Optional date filter for portal scraping
        
    Returns:
        Dictionary with roadmap, changeAnnouncements, and whatsNew data
    """
    # Scrape Entra portal data
    portal_data = await scrape_entra_portal(date_filter)
    
    # Scrape What's New page
    print("\n📚 Scraping Microsoft Learn What's New page...")
    whats_new = scrape_whats_new_page()
    
    return {
        'roadmap': portal_data.get('roadmap'),
        'changeAnnouncements': portal_data.get('changeAnnouncements'),
        'whatsNew': whats_new
    }
