"""
Microsoft Graph API Module

Provides functions for authenticating with and making requests to
the Microsoft Graph API using client credentials flow.

This module handles:
- Token acquisition and caching for client credentials flow
- Building authorization headers for Graph API requests
- Constructing workbook URLs for Excel operations
"""

import os
import logging
import time
from typing import Optional

import httpx

logger = logging.getLogger("mcp-excel-server")

# Microsoft Graph API base URL
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

# Token cache for client credentials flow (for outgoing Graph API requests)
_token_cache = {
    "access_token": None,
    "expires_at": 0,
}


# =============================================================================
# Credential Management
# =============================================================================

def get_client_credentials() -> tuple[str, str, str]:
    """
    Get Azure AD client credentials from environment.
    
    Returns:
        Tuple of (tenant_id, client_id, client_secret)
        
    Raises:
        ValueError: If any required environment variables are missing
    """
    tenant_id = os.getenv("AZURE_TENANT_ID")
    client_id = os.getenv("AZURE_CLIENT_ID")
    client_secret = os.getenv("AZURE_CLIENT_SECRET")
    
    if not all([tenant_id, client_id, client_secret]):
        missing = []
        if not tenant_id:
            missing.append("AZURE_TENANT_ID")
        if not client_id:
            missing.append("AZURE_CLIENT_ID")
        if not client_secret:
            missing.append("AZURE_CLIENT_SECRET")
        raise ValueError(f"Missing required environment variables: {', '.join(missing)}")
    
    return tenant_id, client_id, client_secret


# =============================================================================
# Token Acquisition
# =============================================================================

async def get_access_token() -> str:
    """
    Get Microsoft Graph API access token using client credentials flow.
    
    Implements token caching to avoid unnecessary token requests.
    Tokens are refreshed 5 minutes before expiration.
    
    Returns:
        Valid access token string
        
    Raises:
        ValueError: If token acquisition fails
    """
    global _token_cache
    
    # Check if we have a valid cached token (with 5-minute buffer)
    if _token_cache["access_token"] and time.time() < _token_cache["expires_at"] - 300:
        return _token_cache["access_token"]
    
    logger.info("Acquiring new access token via client credentials flow")
    
    tenant_id, client_id, client_secret = get_client_credentials()
    
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    
    token_data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }
    
    async with httpx.AsyncClient() as client:
        response = await client.post(
            token_url,
            data=token_data,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=30.0,
        )
        
        if response.status_code != 200:
            error_data = response.json() if response.content else {}
            error_desc = error_data.get("error_description", response.text)
            raise ValueError(f"Failed to acquire token: {error_desc}")
        
        token_response = response.json()
        
        # Cache the token
        _token_cache["access_token"] = token_response["access_token"]
        _token_cache["expires_at"] = time.time() + token_response.get("expires_in", 3600)
        
        logger.info("Successfully acquired new access token")
        return _token_cache["access_token"]


async def get_graph_headers() -> dict:
    """
    Get headers for Microsoft Graph API requests.
    
    Returns:
        Dictionary with Authorization and Content-Type headers
    """
    token = await get_access_token()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


# =============================================================================
# URL Building
# =============================================================================

def build_workbook_url(
    drive_id: str,
    item_id: str,
    site_id: Optional[str] = None,
) -> str:
    """
    Build the base URL for workbook operations.
    
    Args:
        drive_id: The ID of the drive containing the Excel file
        item_id: The ID of the Excel file item
        site_id: Optional SharePoint site ID (for SharePoint-hosted files)
    
    Returns:
        Base URL for workbook operations
    """
    if site_id:
        return f"{GRAPH_API_BASE}/sites/{site_id}/drives/{drive_id}/items/{item_id}/workbook"
    return f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}/workbook"
