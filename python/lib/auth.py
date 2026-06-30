"""
Authentication management for Entra Change Tracker
Handles MSAL device code authentication and DefaultAzureCredential (IWA)
"""

import os
from typing import Optional, Dict

import msal
from azure.identity import DefaultAzureCredential
from msal_extensions import (
    PersistedTokenCache,
    FilePersistence,
    build_encrypted_persistence,
)

from .config import get_config

# Global access token
_access_token: Optional[str] = None


def _get_token_cache_path() -> str:
    """Get the path to the token cache file."""
    return os.path.join(os.path.dirname(os.path.dirname(__file__)), ".token-cache.json")


def _build_token_cache() -> PersistedTokenCache:
    """
    Build an OS-encrypted, persistent MSAL token cache.

    Secrets are encrypted at rest using the platform's secure storage via
    msal-extensions: DPAPI on Windows, Keychain on macOS, and libsecret on
    Linux. Persistence is handled transparently by PersistedTokenCache, so no
    manual serialize/deserialize is required.
    """
    cache_path = _get_token_cache_path()

    try:
        persistence = build_encrypted_persistence(cache_path)

        # A cache written by the legacy plaintext implementation (or by a
        # different machine/user) cannot be decrypted here. Detect that case
        # and discard it so a fresh, encrypted cache is created.
        if os.path.exists(cache_path):
            try:
                persistence.load()
            except Exception:
                print("⚠️ Existing token cache could not be decrypted; "
                      "discarding it and re-authenticating.")
                os.remove(cache_path)
    except Exception as err:
        print(f"⚠️ OS-level encryption unavailable for the token cache ({err}); "
              "falling back to an unencrypted file.")
        persistence = FilePersistence(cache_path)

    return PersistedTokenCache(persistence)


def get_graph_access_token_with_iwa(config: Dict[str, str]) -> str:
    """
    Acquire an access token using DefaultAzureCredential (IWA) for Microsoft Graph.
    Uses the currently logged-in user's credentials.
    
    Args:
        config: Configuration dictionary (optional tenantId)
        
    Returns:
        Access token string
    """
    print("🔄 Using DefaultAzureCredential (Integrated Windows Authentication)...")
    
    # Create credential with optional tenant ID
    credential_kwargs = {}
    if config.get('tenantId'):
        credential_kwargs['tenant_id'] = config['tenantId']
    
    credential = DefaultAzureCredential(**credential_kwargs)
    
    # Get token for Microsoft Graph
    token_response = credential.get_token("https://graph.microsoft.com/.default")
    print("✅ Token acquired via DefaultAzureCredential")
    
    return token_response.token


def get_graph_access_token_with_device_code(config: Dict[str, str]) -> str:
    """
    Acquire an access token using device code flow for Microsoft Graph.
    
    Args:
        config: Configuration dictionary with clientId and tenantId
        
    Returns:
        Access token string
    """
    cache = _build_token_cache()
    
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
        return result["access_token"]
    else:
        raise Exception(f"Authentication failed: {result.get('error_description', result)}")


def initialize_authentication() -> None:
    """
    Initialize authentication and acquire access token.
    Requires configuration to be loaded first via config.load_configuration().
    """
    global _access_token
    
    try:
        config = get_config()
        
        if not config:
            print("⚠️ No configuration loaded - authentication skipped")
            return
        
        # Determine authentication method
        sharepoint_config = config.get('sharepoint', {})
        auth_config = sharepoint_config.get('authentication', {})
        auth_method = (auth_config.get('authMethod') or 'devicecode').lower()
        
        if sharepoint_config.get('siteUrl'):
            if auth_method in ['iwa', 'default']:
                print("🔑 Acquiring Graph access token via DefaultAzureCredential (IWA)...")
                _access_token = get_graph_access_token_with_iwa(auth_config)
                print("✅ Access token acquired successfully")
            elif auth_method == 'devicecode':
                device_code_config = auth_config.get('devicecode', {})
                if not device_code_config.get('clientId') or not device_code_config.get('tenantId'):
                    print("❌ clientId and tenantId are required for device code flow")
                    exit(1)
                print("🔑 Acquiring Graph access token via device code flow...")
                _access_token = get_graph_access_token_with_device_code(device_code_config)
                print("✅ Access token acquired successfully")
            else:
                print(f"❌ Invalid authMethod: \"{auth_method}\". Valid options: 'devicecode', 'iwa', 'default'")
                exit(1)
        else:
            print("⚠️ config.json incomplete (siteUrl missing) - data saved locally only")
    
    except Exception as err:
        print(f"⚠️ Error during authentication: {err}")
        raise


def get_access_token() -> Optional[str]:
    """
    Get the access token.
    
    Returns:
        The access token string or None
    """
    return _access_token
