"""
Authentication and configuration management for Entra Change Tracker
Handles MSAL device code authentication and config loading
"""

import json
import os
from pathlib import Path
from typing import Optional, Dict, Any

import msal

VALID_DATE_FILTERS = ["Last 1 month", "Last 3 months", "Last 6 months", "Last 1 year"]

# Global configuration variables
_config: Optional[Dict[str, Any]] = None
_date_filter: Optional[str] = None
_access_token: Optional[str] = None


def _get_token_cache_path() -> str:
    """Get the path to the token cache file."""
    return os.path.join(os.path.dirname(os.path.dirname(__file__)), ".token-cache.json")


def _load_token_cache() -> msal.SerializableTokenCache:
    """Load or create a token cache."""
    cache = msal.SerializableTokenCache()
    cache_path = _get_token_cache_path()
    
    if os.path.exists(cache_path):
        with open(cache_path, 'r') as f:
            cache.deserialize(f.read())
    
    return cache


def _save_token_cache(cache: msal.SerializableTokenCache) -> None:
    """Save the token cache to disk."""
    cache_path = _get_token_cache_path()
    
    if cache.has_state_changed:
        with open(cache_path, 'w') as f:
            f.write(cache.serialize())


def get_graph_access_token_with_device_code(config: Dict[str, str]) -> str:
    """
    Acquire an access token using device code flow for Microsoft Graph.
    
    Args:
        config: Configuration dictionary with clientId and tenantId
        
    Returns:
        Access token string
    """
    cache = _load_token_cache()
    
    app = msal.PublicClientApplication(
        client_id=config['clientId'],
        authority=f"https://login.microsoftonline.com/{config['tenantId']}",
        token_cache=cache
    )
    
    # Try silent token acquisition first
    accounts = app.get_accounts()
    
    if accounts:
        try:
            print("🔄 Attempting to use cached token...")
            result = app.acquire_token_silent(
                scopes=["https://graph.microsoft.com/.default"],
                account=accounts[0]
            )
            if result and "access_token" in result:
                print("✅ Using cached token")
                _save_token_cache(cache)
                return result["access_token"]
        except Exception:
            print("⚠️ Cached token invalid or expired, requesting new token...")
    
    # Device code flow
    flow = app.initiate_device_flow(scopes=["https://graph.microsoft.com/.default"])
    
    if "user_code" not in flow:
        raise ValueError(
            "Failed to create device flow. Error: %s" % json.dumps(flow, indent=4)
        )
    
    print("\n🔐 Device Code Authentication Required")
    print("=" * 60)
    print(flow["message"])
    print("=" * 60)
    
    result = app.acquire_token_by_device_flow(flow)
    
    if "access_token" in result:
        _save_token_cache(cache)
        return result["access_token"]
    else:
        raise Exception(f"Authentication failed: {result.get('error_description', result)}")


def initialize_configuration() -> None:
    """
    Initialize configuration from config.json and acquire access token.
    """
    global _config, _date_filter, _access_token
    
    try:
        config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config.json")
        
        if not os.path.exists(config_path):
            print("⚠️ config.json not found - data will only be saved locally")
            print("📅 No date filter specified, showing all results")
            return
        
        with open(config_path, 'r', encoding='utf-8') as f:
            loadedConfig = json.load(f)
        
        # Validate and set date filter
        if loadedConfig.get('dateFilter'):
            if loadedConfig['dateFilter'] in VALID_DATE_FILTERS:
                _date_filter = loadedConfig['dateFilter']
                print(f"📅 Using date filter from config: {_date_filter}")
            else:
                print(f"❌ Invalid date filter in config: \"{loadedConfig['dateFilter']}\"")
                print(f"   Valid options: {', '.join(VALID_DATE_FILTERS)}")
                exit(1)
        else:
            print("📅 No date filter specified in config, showing all results")
        
        if loadedConfig.get('siteUrl') and loadedConfig.get('clientId') and loadedConfig.get('tenantId'):
            print("🔑 Acquiring Graph access token via device code flow...")
            _access_token = get_graph_access_token_with_device_code(loadedConfig)
            print("✅ Access token acquired successfully")
            
            _config = loadedConfig
            print("✅ SharePoint/Graph configuration loaded from config.json")
        else:
            print("⚠️ config.json incomplete (siteUrl/clientId/tenantId missing) - data saved locally only")
    
    except Exception as err:
        print(f"⚠️ Error loading SharePoint config: {err}")
        _config = None


def get_configuration() -> Dict[str, Any]:
    """
    Get the current configuration.
    
    Returns:
        Dictionary with config, dateFilter, and accessToken
    """
    return {
        'config': _config,
        'dateFilter': _date_filter,
        'accessToken': _access_token
    }
