"""
Configuration management for Entra Change Tracker
Handles loading and validation of config.json
"""

import json
import os
from typing import Optional, Dict, Any

VALID_DATE_FILTERS = ["Last 1 month", "Last 3 months", "Last 6 months", "Last 1 year"]

# Global configuration state
_config: Optional[Dict[str, Any]] = None
_date_filter: Optional[str] = None


def load_configuration() -> None:
    """
    Load and validate configuration from config.json.
    """
    global _config, _date_filter
    
    try:
        config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config.json")
        
        if not os.path.exists(config_path):
            print("⚠️ config.json not found - data will only be saved locally")
            print("📅 No date filter specified, showing all results")
            return
        
        with open(config_path, 'r', encoding='utf-8') as f:
            _config = json.load(f)
        
        # Validate and set date filter
        config_date_filter = _config.get('browserScraping', {}).get('dateFilter')
        if config_date_filter:
            if config_date_filter in VALID_DATE_FILTERS:
                _date_filter = config_date_filter
                print(f"📅 Using date filter from config: {_date_filter}")
            else:
                print(f"❌ Invalid date filter in config: \"{config_date_filter}\"")
                print(f"   Valid options: {', '.join(VALID_DATE_FILTERS)}")
                exit(1)
        else:
            print("📅 No date filter specified in config, showing all results")
    
    except Exception as err:
        print(f"❌ Error loading config.json: {err}")
        exit(1)


def get_config() -> Optional[Dict[str, Any]]:
    """
    Get the loaded configuration.
    
    Returns:
        The configuration dictionary or None if not loaded
    """
    return _config


def get_date_filter() -> Optional[str]:
    """
    Get the date filter from configuration.
    
    Returns:
        The date filter string or None
    """
    return _date_filter


def is_config_loaded() -> bool:
    """
    Check if configuration is loaded.
    
    Returns:
        True if configuration is loaded
    """
    return _config is not None
