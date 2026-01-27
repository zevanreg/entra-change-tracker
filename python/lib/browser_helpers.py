"""
Browser automation helper functions for Entra portal scraping
"""

import json
import os
import re
import time
from datetime import datetime
from typing import Dict, Any, List, Optional, Tuple

from playwright.async_api import Frame, Page

# Load configuration
config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config.json')
try:
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
except Exception as err:
    raise RuntimeError(f"Failed to load config.json from {config_path}: {err}")

# ==================== CONSTANTS ====================

TIMEOUTS = {
    'SPLASH_SCREEN': config['timeouts']['splashScreen'],
    'PROGRESS_DOTS': config['timeouts']['progressDots'],
    'GENERAL_WAIT': config['timeouts']['generalWait'],
    'CLICK': config['timeouts']['click'],
    'DETACH': config['timeouts']['detach'],
    'CLOSE_PANE': config['timeouts']['closePane'],
    'BUTTON_CLOSE': config['timeouts']['buttonClose'],
    'SHORT_DELAY': config['timeouts']['shortDelay'],
    'MENU_DELAY': config['timeouts']['menuDelay'],
    'CHECKBOX_DELAY': config['timeouts']['checkboxDelay']
}

SELECTORS = {
    'SPLASH_SCREEN': config['selectors']['splashScreen'],
    'DETAILS_IFRAME': config['selectors']['detailsIframe'],
    'DETAILS_ROW': config['selectors']['detailsRow'],
    'DETAILS_ROW_CHECK': config['selectors']['detailsRowCheck'],
    'DETAILS_ROW_FIELDS': config['selectors']['detailsRowFields'],
    'DETAILS_ROW_CELL': config['selectors']['detailsRowCell'],
    'PROGRESS_DOTS': config['selectors']['progressDots'],
    'CLOSE_BUTTON': config['selectors']['closeButton'],
    'SCROLLABLE_CONTAINER': config['selectors']['scrollableContainer'],
    'FILTER_BUTTON_CONTAINER': config['selectors']['filterButtonContainer'],
    'APPLY_BUTTON': config['selectors']['applyButton'],
    'RADIO_LABEL': config['selectors']['radioLabel']
}

SCRAPER_CONFIG = {
    'SCROLL_STEP_PX': config['scraperConfig']['scrollStepPx'],
    'PASS_DELAY_MS': config['scraperConfig']['passDelayMs'],
    'MAX_IDLE_PASSES': config['scraperConfig']['maxIdlePasses'],
    'TOTAL_TIMEOUT_MS': config['scraperConfig']['totalTimeoutMs'],
    'MAX_RETRY_ATTEMPTS': config['scraperConfig']['maxRetryAttempts'],
    'CLICK_OUTSIDE_COORDS': config['scraperConfig']['clickOutsideCoords']
}

TEXT_PATTERNS = {
    'OVERVIEW': config['textPatterns']['overview'],
    'NEXT_STEPS': config['textPatterns']['nextSteps'],
    'WHAT_IS_CHANGING': config['textPatterns']['whatIsChanging'],
    'ROADMAP_DESCRIPTION': config['textPatterns']['roadmapDescription']
}


# ==================== HELPER FUNCTIONS ====================

async def wait_for_splash_screen(frame: Frame, timeout: int = None) -> None:
    """
    Waits for the Entra splash screen to disappear.
    
    Args:
        frame: The frame to check
        timeout: Maximum wait time in ms
    """
    if timeout is None:
        timeout = TIMEOUTS['SPLASH_SCREEN']
    
    try:
        splash_screen = frame.locator(SELECTORS['SPLASH_SCREEN'])
        await splash_screen.wait_for(state='hidden', timeout=timeout)
    except Exception:
        # Splash screen might not exist or already hidden
        print('Splash screen not found or already hidden')


async def click_tab(frame: Frame, tab_name: str) -> bool:
    """
    Clicks on a tab within a frame using role-based locators.
    
    Args:
        frame: The frame containing the tab
        tab_name: The name of the tab to click (string or regex pattern)
        
    Returns:
        True if the tab was found and clicked, false otherwise
    """
    try:
        # Wait for splash screen to disappear first
        await wait_for_splash_screen(frame)
        
        # Handle regex patterns
        if isinstance(tab_name, str) and tab_name.startswith('/') and tab_name.endswith('/i'):
            # Convert JS regex to Python re pattern
            pattern = re.compile(tab_name[1:-2], re.IGNORECASE)
            role_tab = frame.get_by_role('tab', name=pattern)
        else:
            role_tab = frame.get_by_role('tab', name=tab_name)
        
        count = await role_tab.count()
        if count > 0:
            # Wait for tab to be visible and stable
            await role_tab.first.wait_for(state='visible', timeout=TIMEOUTS['GENERAL_WAIT'])
            
            try:
                # Try normal click first
                await role_tab.first.click(timeout=TIMEOUTS['CLICK'])
            except Exception as click_err:
                # If intercepted, force click
                print(f"Normal click failed, forcing click on tab \"{tab_name}\"")
                await role_tab.first.click(force=True, timeout=TIMEOUTS['CLICK'])
            
            # Wait a moment for tab content to load
            await frame.wait_for_timeout(TIMEOUTS['SHORT_DELAY'])
            return True
    except Exception as err:
        print(f"Error clicking tab \"{tab_name}\": {err}")
    
    return False


async def set_date_range_filter(frame: Frame, filter_option: str) -> bool:
    """
    Sets the date range filter to a specific option.
    
    Args:
        frame: The frame containing the filter
        filter_option: The text of the filter option to select (e.g., "Last 1 month")
        
    Returns:
        True if the filter was set successfully
    """
    try:
        # Click the filter button - find button inside div with data-selection-index='1'
        filter_button = frame.locator(f"{SELECTORS['FILTER_BUTTON_CONTAINER']} button").first
        await filter_button.wait_for(state='visible', timeout=TIMEOUTS['GENERAL_WAIT'])
        await filter_button.click(timeout=TIMEOUTS['CLICK'])
        print('Filter button clicked')
        
        # Wait for the filter menu to appear
        await frame.wait_for_timeout(TIMEOUTS['MENU_DELAY'])
        
        # Find and click the radio input with the matching label
        radio_label = frame.locator(f"span{SELECTORS['RADIO_LABEL']}:has-text(\"{filter_option}\")")
        await radio_label.wait_for(state='visible', timeout=TIMEOUTS['CLICK'])
        
        # Click the label to select the radio button
        await radio_label.click(timeout=TIMEOUTS['CLICK'])
        print(f"Selected filter option: {filter_option}")
        
        # Wait a moment for the selection to register
        await frame.wait_for_timeout(TIMEOUTS['MENU_DELAY'])
        
        # Click the Apply button
        apply_button = frame.locator(SELECTORS['APPLY_BUTTON'])
        if await apply_button.count() > 0 and not await apply_button.is_disabled():
            await apply_button.click(timeout=TIMEOUTS['CLICK'])
            print('Apply button clicked')
            await frame.wait_for_timeout(TIMEOUTS['SHORT_DELAY'])
        
        return True
    except Exception as err:
        print(f"Error setting date range filter to \"{filter_option}\": {err}")
        return False


# ==================== DETAILS EXTRACTION HELPERS ====================

async def extract_overview(details_frame: Frame, row_index: int) -> str:
    """
    Extracts overview text from the details pane.
    
    Args:
        details_frame: The details iframe
        row_index: Row index for logging
        
    Returns:
        The extracted overview text
    """
    try:
        if details_frame.is_detached():
            return ''
        
        overview_h3 = details_frame.locator('h3').filter(has_text=TEXT_PATTERNS['OVERVIEW'])
        
        if await overview_h3.count() > 0:
            parent = overview_h3.locator('..')
            spans = parent.locator('span')
            
            if await spans.count() > 0:
                text = await spans.first.inner_text()
                return text if text else ''
    except Exception as err:
        print(f"Could not extract overview for row {row_index}: {err}")
    
    return ''


async def extract_url(details_frame: Frame, row_index: int) -> str:
    """
    Extracts URL from the "Next steps" section.
    
    Args:
        details_frame: The details iframe
        row_index: Row index for logging
        
    Returns:
        The extracted URL
    """
    try:
        if details_frame.is_detached():
            return ''
        
        next_steps_h3 = details_frame.locator('h3').filter(has_text=TEXT_PATTERNS['NEXT_STEPS'])
        
        if await next_steps_h3.count() > 0:
            parent = next_steps_h3.locator('..')
            links = parent.locator('a')
            
            if await links.count() > 0:
                href = await links.first.get_attribute('href')
                return href if href else ''
    except Exception as err:
        print(f"Could not extract URL for row {row_index}: {err}")
    
    return ''


async def extract_description(details_frame: Frame, row_index: int) -> str:
    """
    Extracts description text from either "What is changing" or roadmap section.
    
    Args:
        details_frame: The details iframe
        row_index: Row index for logging
        
    Returns:
        The extracted description text
    """
    try:
        if details_frame.is_detached():
            return ''
        
        # Try "What is changing" first (change announcements)
        desc_h3 = details_frame.locator('h3').filter(has_text=TEXT_PATTERNS['WHAT_IS_CHANGING'])
        
        if await desc_h3.count() > 0:
            # Get the span that is the next sibling of the h3
            next_span = desc_h3.locator('xpath=following-sibling::span[1]')
            
            if await next_span.count() > 0:
                text = await next_span.inner_text()
                return text if text else ''
        else:
            # Fall back to "Here's what you will see in this release:" (roadmap)
            desc_h3 = details_frame.locator('h3').filter(has_text=TEXT_PATTERNS['ROADMAP_DESCRIPTION'])
            
            if await desc_h3.count() > 0:
                parent = desc_h3.locator('..')
                paragraphs = parent.locator('p')
                
                if await paragraphs.count() > 0:
                    text = await paragraphs.first.inner_text()
                    return text if text else ''
    except Exception as err:
        print(f"Could not extract description for row {row_index}: {err}")
    
    return ''


# ==================== IFRAME MANAGEMENT ====================

async def open_details_pane(page: Page, frame: Frame, row_index: int) -> None:
    """
    Opens the details pane by clicking a row's checkbox.
    
    Args:
        page: The main page
        frame: The frame containing the row
        row_index: The index of the row to click
    """
    row = frame.locator(f"{SELECTORS['DETAILS_ROW']}[data-item-index='{row_index}']").first
    checkbox = row.locator(SELECTORS['DETAILS_ROW_CHECK'])
    
    # Check if the row is already selected (checked)
    is_checked = await checkbox.get_attribute('aria-checked')
    if is_checked == 'true':
        # Uncheck it first by clicking
        await checkbox.click(timeout=TIMEOUTS['CLICK'])
        await page.wait_for_timeout(TIMEOUTS['CHECKBOX_DELAY'])
    
    # Now click to select and open details pane
    await checkbox.click(timeout=TIMEOUTS['CLICK'])
    
    # Wait for at least one details iframe to appear
    await page.locator(f"{SELECTORS['DETAILS_IFRAME']}:visible").first.wait_for(
        state='attached',
        timeout=TIMEOUTS['GENERAL_WAIT']
    )


async def find_correct_iframe(page: Page, row_title: str, row_index: int) -> Optional[Frame]:
    """
    Finds the correct details iframe by matching the row title.
    
    Args:
        page: The main page
        row_title: The title to match
        row_index: Row index for logging
        
    Returns:
        The matched iframe or None
    """
    all_iframes = await page.locator(SELECTORS['DETAILS_IFRAME']).all()
    
    for iframe in all_iframes:
        try:
            await iframe.wait_for(state='attached', timeout=TIMEOUTS['GENERAL_WAIT'])
            
            iframe_handle = await iframe.element_handle()
            frame = await iframe_handle.content_frame()
            
            if not frame:
                continue
            
            # Wait for progress dots to disappear
            try:
                progress_dots = frame.locator(SELECTORS['PROGRESS_DOTS'])
                await progress_dots.wait_for(state='hidden', timeout=TIMEOUTS['PROGRESS_DOTS'])
            except Exception:
                # Progress dots might not exist or already hidden
                pass
            
            await frame.wait_for_load_state('domcontentloaded')
            
            if row_title:
                # Check if this iframe contains an h3 with the row title
                title_h3 = frame.locator('h3').filter(has_text=row_title)
                if await title_h3.count() > 0:
                    return frame
            else:
                # If we couldn't get the title, use this iframe
                return frame
        except Exception:
            # Skip this iframe if we can't access it
            continue
    
    print(f"Could not find iframe with matching title for row {row_index}")
    return None


async def close_details_pane(page: Page, row_index: int) -> None:
    """
    Closes the details pane and waits for iframe removal.
    
    Args:
        page: The main page
        row_index: Row index for logging
    """
    try:
        # Method 1: Look for close button in the main page (Azure blade close button)
        close_button = page.locator(SELECTORS['CLOSE_BUTTON']).last
        if await close_button.count() > 0 and await close_button.is_visible():
            await close_button.click(timeout=TIMEOUTS['BUTTON_CLOSE'])
        else:
            # Method 2: Click outside the iframe to close it
            await page.mouse.click(
                SCRAPER_CONFIG['CLICK_OUTSIDE_COORDS']['x'],
                SCRAPER_CONFIG['CLICK_OUTSIDE_COORDS']['y']
            )
        
        # Wait for all details iframes to be removed from DOM
        await page.wait_for_function(
            "selector => document.querySelectorAll(selector).length === 0",
            arg=SELECTORS['DETAILS_IFRAME'],
            timeout=TIMEOUTS['CLOSE_PANE']
        )
    except Exception as err:
        # If close button/click didn't work, try ESC key as fallback
        print(f"Could not close details pane with button/click for row {row_index}, trying ESC key... {err}")
        try:
            await page.keyboard.press('Escape')
            await page.wait_for_function(
                "selector => document.querySelectorAll(selector).length === 0",
                arg=SELECTORS['DETAILS_IFRAME'],
                timeout=TIMEOUTS['CLOSE_PANE']
            )
        except Exception as esc_err:
            print(f"Details iframe did not detach for row {row_index}, continuing anyway...")


# ==================== MAIN EXTRACTION FUNCTION ====================

async def extract_row_details(
    page: Page,
    frame: Frame,
    row_index: int,
    row_title: str = ''
) -> Dict[str, str]:
    """
    Extracts details from a row by clicking it and scraping the details pane.
    
    Args:
        page: The main page
        frame: The frame containing the row
        row_index: The index of the row to click
        row_title: The title of the row (for iframe matching)
        
    Returns:
        Dictionary with url, description, and overview
    """
    start_time = time.time()
    print(f"[{datetime.now().isoformat()}] 🔵 START extracting details for row {row_index}")
    
    try:
        # Open the details pane
        await open_details_pane(page, frame, row_index)
        
        # Find the correct iframe
        details_frame = await find_correct_iframe(page, row_title, row_index)
        
        if not details_frame:
            print(f"Details frame not available for row {row_index}")
            return {'url': '', 'description': '', 'overview': ''}
        
        # Extract all details using helper functions
        overview = await extract_overview(details_frame, row_index)
        url = await extract_url(details_frame, row_index)
        description = await extract_description(details_frame, row_index)
        
        # Close the details pane
        await close_details_pane(page, row_index)
        
        elapsed = (time.time() - start_time) * 1000
        print(f"[{datetime.now().isoformat()}] ✅ END extracting details for row {row_index} (took {elapsed:.0f}ms)")
        
        return {
            'url': url.strip(),
            'description': description.strip(),
            'overview': overview.strip()
        }
    except Exception as err:
        elapsed = (time.time() - start_time) * 1000
        print(f"[{datetime.now().isoformat()}] ❌ ERROR extracting details for row {row_index} (took {elapsed:.0f}ms): {err}")
        
        # Attempt to close any open pane
        try:
            details_iframe_locator = page.locator(SELECTORS['DETAILS_IFRAME']).last
            
            # Try close button first
            close_button = page.locator('button[aria-label="Close"]').last
            if await close_button.count() > 0:
                await close_button.click(timeout=TIMEOUTS['CLOSE_PANE'])
            else:
                # Try clicking outside
                await page.mouse.click(
                    SCRAPER_CONFIG['CLICK_OUTSIDE_COORDS']['x'],
                    SCRAPER_CONFIG['CLICK_OUTSIDE_COORDS']['y']
                )
            
            await details_iframe_locator.wait_for(state='detached', timeout=TIMEOUTS['DETACH'])
        except Exception:
            pass
        
        return {'url': '', 'description': '', 'overview': ''}


# ==================== LIST SCRAPING ====================

async def scrape_details_list(page: Page, frame: Frame, extract_details: bool = True) -> List[Dict[str, Any]]:
    """
    Scrapes all rows from a Fluent UI DetailsList with virtualized scrolling.
    
    Args:
        page: The main page
        frame: The frame containing the DetailsList
        extract_details: Whether to extract details (URL, description, overview) by clicking each row
        
    Returns:
        Array of row objects with field names and details
    """
    # Accumulator for processed rows
    processed_rows = []
    seen_indices = set()
    
    idle_passes = 0
    start_time = time.time()
    
    # Scroll to top first
    await frame.evaluate(f"""
        (selector) => {{
            const container = document.querySelector(selector) ||
                              document.querySelector(".ms-DetailsList") ||
                              document.querySelector(".ms-List") ||
                              document.scrollingElement ||
                              document.documentElement;
            container.scrollTop = 0;
            window.scrollTo(0, 0);
        }}
    """, SELECTORS['SCROLLABLE_CONTAINER'])
    await page.wait_for_timeout(SCRAPER_CONFIG['PASS_DELAY_MS'])
    
    while True:
        # Timeout guard
        if (time.time() - start_time) * 1000 > SCRAPER_CONFIG['TOTAL_TIMEOUT_MS']:
            print("Timeout reached; returning accumulated rows.")
            break
        
        # Get currently visible rows
        visible_rows = await frame.locator(SELECTORS['DETAILS_ROW']).all()
        new_rows_found = False
        
        for row in visible_rows:
            # Get row index
            idx = await row.get_attribute('data-item-index')
            if not idx:
                row_id = await row.get_attribute('id')
                if row_id:
                    m = re.search(r'-(\d+)$', row_id)
                    if m:
                        idx = m.group(1)
            
            index_num = int(idx) if idx and idx.isdigit() else None
            if index_num is None or index_num in seen_indices:
                continue
            
            # Mark as seen
            seen_indices.add(index_num)
            new_rows_found = True
            
            # Extract basic row data
            field_container = row.locator(SELECTORS['DETAILS_ROW_FIELDS'])
            cells = await field_container.locator(SELECTORS['DETAILS_ROW_CELL']).all()
            
            obj = {}
            for i, cell in enumerate(cells):
                key = await cell.get_attribute("data-automation-key")
                if not key:
                    key = f"col{i}"
                value = await cell.inner_text()
                obj[key] = value.strip()
            
            # Extract details by clicking the row (with retry logic for empty descriptions)
            details = {'url': '', 'description': '', 'overview': ''}
            
            if extract_details:
                for attempt in range(1, SCRAPER_CONFIG['MAX_RETRY_ATTEMPTS'] + 1):
                    details = await extract_row_details(page, frame, index_num, obj.get('title', ''))
                    
                    if details['description']:
                        # Description found, break out of retry loop
                        break
                    
                    if attempt < SCRAPER_CONFIG['MAX_RETRY_ATTEMPTS']:
                        print(f"Empty description for row {index_num}, retrying (attempt {attempt}/{SCRAPER_CONFIG['MAX_RETRY_ATTEMPTS']})...")
                    else:
                        print(f"Empty description for row {index_num} after {SCRAPER_CONFIG['MAX_RETRY_ATTEMPTS']} attempts")
            
            obj['url'] = details['url']
            obj['description'] = details['description']
            obj['overview'] = details['overview']
            
            processed_rows.append(obj)
        
        if new_rows_found:
            idle_passes = 0
        else:
            idle_passes += 1
        
        # Check if at bottom
        scroll_info = await frame.evaluate(f"""
            (selector) => {{
                const container = document.querySelector(selector) ||
                                  document.querySelector(".ms-DetailsList") ||
                                  document.querySelector(".ms-List") ||
                                  document.scrollingElement ||
                                  document.documentElement;
                const maxScroll = Math.max(0, container.scrollHeight - container.clientHeight);
                const atBottom = container.scrollTop >= maxScroll - 2;
                return {{ atBottom: atBottom, currentScroll: container.scrollTop, maxScroll: maxScroll }};
            }}
        """, SELECTORS['SCROLLABLE_CONTAINER'])
        
        if scroll_info['atBottom'] and idle_passes >= SCRAPER_CONFIG['MAX_IDLE_PASSES']:
            print("Reached bottom and stabilized.")
            break
        
        # Scroll down
        await frame.evaluate(f"""
            ({{ selector, step }}) => {{
                const container = document.querySelector(selector) ||
                                  document.querySelector(".ms-DetailsList") ||
                                  document.querySelector(".ms-List") ||
                                  document.scrollingElement ||
                                  document.documentElement;
                const maxScroll = Math.max(0, container.scrollHeight - container.clientHeight);
                container.scrollTop = Math.min(container.scrollTop + step, maxScroll);
                window.scrollTo(0, container.scrollTop);
            }}
        """, {'selector': SELECTORS['SCROLLABLE_CONTAINER'], 'step': SCRAPER_CONFIG['SCROLL_STEP_PX']})
        
        await page.wait_for_timeout(SCRAPER_CONFIG['PASS_DELAY_MS'])
    
    print(f"✅ Collected and processed {len(processed_rows)} unique rows with details")
    return processed_rows
