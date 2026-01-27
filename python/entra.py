"""
Entra Change Tracker - Main Script
Orchestrates scraping and data insertion into SharePoint
"""

import asyncio
import json
import os
from datetime import datetime
from pathlib import Path

from lib.config import load_configuration, get_config, get_date_filter
from lib.auth import initialize_authentication, get_access_token
from lib.sharepoint import insert_into_sharepoint_list, reset_caches
from lib.scraper import scrape_all_sources


def save_to_file(filename: str, data: list, timestamp: str) -> None:
    """
    Save data to JSON file.
    
    Args:
        filename: Base filename
        data: Data to save
        timestamp: Timestamp for filename
    """
    file_path = os.path.join(os.path.dirname(__file__), f"{filename}-{timestamp}.json")
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"💾 Saved {filename} to {file_path}")


async def main():
    """Main execution function."""
    try:
        # Initialize configuration and authenticate
        load_configuration()
        initialize_authentication()
        reset_caches()
        
        config = get_config()
        access_token = get_access_token()
        date_filter = get_date_filter()
        
        # Generate timestamp for file names
        timestamp = datetime.now().isoformat().replace(':', '-').replace('.', '-')[:19]
        
        # Scrape data from all sources
        result = await scrape_all_sources(date_filter)
        roadmap = result['roadmap']
        change_announcements = result['changeAnnouncements']
        whats_new = result['whatsNew']
        
        # Process Roadmap data
        if roadmap:
            if config['browserScraping']['roadmap']['saveToFile']:
                save_to_file('roadmap', roadmap, timestamp)
            
            browser_scraping = config['browserScraping']
            roadmap_config = browser_scraping['roadmap']
            sharepoint_list = roadmap_config['sharepointList']
            roadmap_list_name = sharepoint_list['name']
            
            insert_into_sharepoint_list(roadmap_list_name, roadmap, access_token, config)
        
        # Process Change Announcements data
        if change_announcements:
            if config['browserScraping']['changeAnnouncements']['saveToFile']:
                save_to_file('change-announcements', change_announcements, timestamp)
            
            browser_scraping = config['browserScraping']
            change_config = browser_scraping['changeAnnouncements']
            sharepoint_list = change_config['sharepointList']
            change_announcements_list_name = sharepoint_list['name']
            
            insert_into_sharepoint_list(
                change_announcements_list_name,
                change_announcements,
                access_token,
                config
            )
        
        # Process What's New data
        if whats_new:
            if config['httpScraping']['saveToFile']:
                save_to_file('whats-new', whats_new, timestamp)
            
            http_scraping = config['httpScraping']
            sharepoint_lists = http_scraping['sharepointList']
            whats_new_config = sharepoint_lists['whatsNew']
            whats_new_list_name = whats_new_config['name']
            
            insert_into_sharepoint_list(
                whats_new_list_name,
                whats_new,
                access_token,
                config
            )
        
        print('✅ Script completed successfully.')
    except Exception as error:
        print(f'❌ Error during execution: {error}')
        exit(1)


if __name__ == '__main__':
    asyncio.run(main())
