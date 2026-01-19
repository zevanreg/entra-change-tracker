"""
Entra Change Tracker - Main Script
Orchestrates scraping and data insertion into SharePoint
"""

import asyncio
import json
import os
from datetime import datetime
from pathlib import Path

from lib.auth import initialize_configuration, get_configuration
from lib.sharepoint import insert_into_sharepoint_list, reset_caches
from lib.scraper import scrape_entra_portal


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
        initialize_configuration()
        reset_caches()
        
        app_config = get_configuration()
        config = app_config['config'] if 'config' in app_config else app_config.get('sharepointConfig')
        date_filter = app_config['dateFilter']
        access_token = app_config['accessToken']
        save_to_file_enabled = False
        if config and 'saveToFile' in config:
            save_to_file_enabled = bool(config['saveToFile'])
        
        # Generate timestamp for file names
        timestamp = datetime.now().isoformat().replace(':', '-').replace('.', '-')[:19]
        
        # Scrape data from Entra portal
        result = await scrape_entra_portal(date_filter)
        roadmap = result['roadmap']
        change_announcements = result['changeAnnouncements']
        
        # Process Roadmap data
        if roadmap:
            if save_to_file_enabled:
                save_to_file('roadmap', roadmap, timestamp)
            
            roadmap_list_name = 'EntraRoadmapItems'
            if config and config.get('lists', {}).get('roadmap'):
                roadmap_entry = config['lists']['roadmap']
                roadmap_list_name = roadmap_entry.get('name') if isinstance(roadmap_entry, dict) else roadmap_entry
            
            insert_into_sharepoint_list(roadmap_list_name, roadmap, access_token, config)
        
        # Process Change Announcements data
        if change_announcements:
            if save_to_file_enabled:
                save_to_file('change-announcements', change_announcements, timestamp)
            
            change_announcements_list_name = 'EntraChangeAnnouncements'
            if config and config.get('lists', {}).get('changeAnnouncements'):
                change_entry = config['lists']['changeAnnouncements']
                change_announcements_list_name = change_entry.get('name') if isinstance(change_entry, dict) else change_entry
            
            insert_into_sharepoint_list(
                change_announcements_list_name,
                change_announcements,
                access_token,
                config
            )
        
        print('✅ Script completed successfully.')
    except Exception as error:
        print(f'❌ Error during execution: {error}')
        exit(1)


if __name__ == '__main__':
    asyncio.run(main())
