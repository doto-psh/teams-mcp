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
            "response_mode": "query",
            "prompt": "select_account"  # 항상 계정 선택 화면 표시
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

@server.tool()
async def logout_teams_auth(user_email: str) -> str:
    """
    Logout from Microsoft Teams by clearing stored credentials and session.
    
    Args:
        user_email (str): The user's email address to logout.
        
    Returns:
        str: Logout confirmation message.
    """
    logger.info(f"[logout_teams_auth] Logging out user: {user_email}")

    try:
        from auth.oauth21_session_store import get_oauth21_session_store
        from auth.teams_auth import DEFAULT_CREDENTIALS_DIR
        import os
        import json
        
        store = get_oauth21_session_store()
        
        # Track what was cleared
        cleared_items = []
        
        # 1. Clear from OAuth session store
        with store._lock:
            if user_email in store._sessions:
                session_info = store._sessions[user_email]
                mcp_session_id = session_info.get("mcp_session_id")
                oauth_session_id = session_info.get("session_id")
                
                # Remove from sessions
                del store._sessions[user_email]
                cleared_items.append("OAuth session store")
                
                # Remove from session mappings
                if mcp_session_id and mcp_session_id in store._mcp_session_mapping:
                    del store._mcp_session_mapping[mcp_session_id]
                    cleared_items.append("MCP session mapping")
                
                # Remove from auth bindings
                if mcp_session_id and mcp_session_id in store._session_auth_binding:
                    del store._session_auth_binding[mcp_session_id]
                    cleared_items.append("MCP session binding")
                
                if oauth_session_id and oauth_session_id in store._session_auth_binding:
                    del store._session_auth_binding[oauth_session_id]
                    cleared_items.append("OAuth session binding")
        
        # 2. Clear from persistent credentials directory
        credentials_dir = DEFAULT_CREDENTIALS_DIR
        if credentials_dir and os.path.exists(credentials_dir):
            for filename in os.listdir(credentials_dir):
                if filename.endswith(".json"):
                    filepath = os.path.join(credentials_dir, filename)
                    try:
                        with open(filepath, "r") as f:
                            creds_data = json.load(f)
                        
                        # Check if this credentials file belongs to the user
                        stored_email = creds_data.get("user_email") or creds_data.get("email")
                        if stored_email == user_email:
                            os.remove(filepath)
                            cleared_items.append(f"Persistent credentials ({filename})")
                            logger.info(f"Removed credentials file: {filepath}")
                    except (IOError, json.JSONDecodeError) as e:
                        logger.warning(f"Could not process credentials file {filepath}: {e}")
        
        if cleared_items:
            cleared_list = "\n".join([f"• {item}" for item in cleared_items])
            return f"""✅ Successfully logged out {user_email}

            Cleared:
            {cleared_list}

            You can now authenticate as a different user or re-authenticate with the same account using the start_teams_auth tool."""
        else:
            return f"ℹ️ No active session found for {user_email}. User was already logged out."
    
    except Exception as e:
        logger.error(f"[logout_teams] Error during logout: {e}")
        return f"❌ Error during logout: {str(e)}"
