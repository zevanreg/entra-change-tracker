"""
Entra Portal Scraper
Handles all browser automation and data extraction from Entra portal
"""

import re
from pathlib import Path
from typing import Dict, Any, List, Optional

from playwright.async_api import async_playwright, Page, Frame

from .browser_helpers import (
    wait_for_splash_screen,
    click_tab,
    set_date_range_filter,
    scrape_details_list
)

ENTRA_URL = "https://entra.microsoft.com/#blade/Microsoft_AAD_IAM/ChangeManagementHubList.ReactView"


async def scrape_tab(
    page: Page,
    frame: Frame,
    tab_name: str,
    date_filter: Optional[str] = None
) -> Optional[List[Dict[str, Any]]]:
    """
    Scrape data from a specific tab.
    
    Args:
        page: The page
        frame: The iframe
        tab_name: The name of the tab to scrape
        date_filter: Optional date filter to apply
        
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
    data = await scrape_details_list(page, frame)
    print(f"✅ Extracted {len(data)} items from {tab_name}")
    
    return data


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
            await page.goto(ENTRA_URL, wait_until='domcontentloaded', timeout=60000)

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

            # Scrape Roadmap
            roadmap = await scrape_tab(page, frame, '/^Roadmap$/i', date_filter)

            # Scrape Change Announcements
            change_announcements = await scrape_tab(page, frame, '/^Change announcements$/i', date_filter)

            return {
                'roadmap': roadmap,
                'changeAnnouncements': change_announcements
            }
        finally:
            await context.close()
            print('✅ Browser closed.')
