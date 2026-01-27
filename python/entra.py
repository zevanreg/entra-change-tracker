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
        
        # Generate timestamp for file names
        timestamp = datetime.now().isoformat().replace(':', '-').replace('.', '-')[:19]
        
        # Scrape data from all sources
        result = await scrape_all_sources()
        roadmap = result['roadmap']
        change_announcements = result['changeAnnouncements']
        whats_new = result['whatsNew']
        
        # Process Roadmap data
        if roadmap:
            if config['browserScraping']['roadmap']['saveToFile']:
                save_to_file('roadmap', roadmap, timestamp)
            
            roadmap_list_name = config['browserScraping']['roadmap']['sharepointList']['name']
            insert_into_sharepoint_list(roadmap_list_name, roadmap)
        
        # Process Change Announcements data
        if change_announcements:
            if config['browserScraping']['changeAnnouncements']['saveToFile']:
                save_to_file('change-announcements', change_announcements, timestamp)
            
            change_announcements_list_name = config['browserScraping']['changeAnnouncements']['sharepointList']['name']
            insert_into_sharepoint_list(change_announcements_list_name, change_announcements)
        
        # Process What's New data
        if whats_new:
            if config['httpScraping']['saveToFile']:
                save_to_file('whats-new', whats_new, timestamp)
            
            whats_new_list_name = config['httpScraping']['sharepointList']['whatsNew']['name']
            insert_into_sharepoint_list(whats_new_list_name, whats_new)
        
        print('✅ Script completed successfully.')
    except Exception as error:
        print(f'❌ Error during execution: {error}')
        exit(1)


if __name__ == '__main__':
    asyncio.run(main())
