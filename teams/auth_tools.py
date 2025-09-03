"""
Microsoft Teams Tools

This module provides MCP tools for interacting with Microsoft Teams via Graph API.
"""

import json
import logging
from typing import List, Dict, Any
from auth.service_decorator_teams import require_teams_service
from core.server import server

logger = logging.getLogger(__name__)

@server.tool()
async def start_teams_auth(user_email: str) -> str:
    """
    Initiate Microsoft Teams OAuth authentication flow.
    
    Args:
        user_email (str): The user's email address for authentication.
        
    Returns:
        str: Authentication instructions.
    """
    logger.info(f"[start_teams_auth] Starting auth flow for: {user_email}")
    
    try:
        from auth.oauth_config import get_oauth_config
        from auth.scopes import get_current_scopes
        import secrets
        import urllib.parse
        
        config = get_oauth_config()
        
        if not config.is_configured():
            return "❌ Microsoft OAuth credentials not configured. Please set MICROSOFT_OAUTH_CLIENT_ID and MICROSOFT_OAUTH_CLIENT_SECRET environment variables."
        
        # Debug information
        logger.debug(f"Client ID: {config.client_id}")
        logger.debug(f"Redirect URI: {config.redirect_uri}")
        logger.debug(f"Tenant ID: {config.tenant_id}")
        
        # Generate state parameter for security
        state = secrets.token_urlsafe(32)
        scopes = get_current_scopes()
        
        # Simplified scope format (remove full URLs)
        simplified_scopes = []
        for scope in scopes:
            if scope.startswith("https://graph.microsoft.com/"):
                simplified_scopes.append(scope.replace("https://graph.microsoft.com/", ""))
            else:
                simplified_scopes.append(scope)
        
        # Manual URL construction for debugging
        base_auth_url = f"https://login.microsoftonline.com/{config.tenant_id}/oauth2/v2.0/authorize"
        
        # URL encode parameters properly
        params = {
            "client_id": config.client_id,
            "response_type": "code",
            "redirect_uri": config.redirect_uri,
            "scope": " ".join(simplified_scopes),
            "state": state,
            "response_mode": "query"
        }
        
        # Build query string manually
        query_parts = []
        for key, value in params.items():
            encoded_value = urllib.parse.quote_plus(str(value))
            query_parts.append(f"{key}={encoded_value}")
        
        query_string = "&".join(query_parts)
        auth_url = f"{base_auth_url}?{query_string}"
        
        return f"""🔐 Microsoft Teams Authentication Required

Debug Information:
• Client ID: {config.client_id}
• Tenant: {config.tenant_id}
• Redirect URI: {config.redirect_uri}
• Scopes: {', '.join(simplified_scopes)}

Please visit the following URL to authenticate:

{auth_url}

After authentication, you'll be redirected back to complete the setup.
"""
    
    except Exception as e:
        logger.error(f"[start_teams_auth] Error: {e}")
        return f"❌ Error starting authentication: {str(e)}"