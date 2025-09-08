"""
Microsoft Teams Users Tools

This module provides MCP tools for user management and search via Microsoft Graph API.
"""

import json
import logging
from typing import Dict, Any, Optional
from auth.service_decorator_teams import require_teams_service
from core.server import server

logger = logging.getLogger(__name__)

@server.tool()
@require_teams_service("teams", "teams_read")
async def get_current_user(service, user_email: str) -> str:
    """
    Get the current authenticated user's profile information including display name, email, job title, and department.
    
    Args:
        user_email (str): The user's email address. Required.
        
    Returns:
        str: JSON string containing current user's profile information.
    """
    logger.info(f"[get_current_user] Fetching current user profile for: {user_email}")
    
    try:
        user = await service.get("/me")
        
        user_summary = {
            "displayName": user.get("displayName"),
            "userPrincipalName": user.get("userPrincipalName"),
            "mail": user.get("mail"),
            "id": user.get("id"),
            "jobTitle": user.get("jobTitle"),
            "department": user.get("department")
        }
        
        return json.dumps(user_summary, indent=2)
        
    except Exception as e:
        logger.error(f"[get_current_user] Error: {e}")
        return f"❌ Error: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def search_users(service, user_email: str, query: str) -> str:
    """
    Search for users in the organization by name or email address. Returns matching users with their basic profile information.
    
    Args:
        user_email (str): The user's email address. Required.
        query (str): Search query (name or email)
        
    Returns:
        str: JSON string containing matching users.
    """
    logger.info(f"[search_users] Searching users with query '{query}', user: {user_email}")
    
    try:
        # Build filter query for searching users
        # Microsoft Graph supports startswith function for filtering
        filter_query = (
            f"startswith(displayName,'{query}') or "
            f"startswith(mail,'{query}') or "
            f"startswith(userPrincipalName,'{query}')"
        )
        
        response = await service.get(f"/users?$filter={filter_query}")
        
        if not response.get("value"):
            return json.dumps({"message": "No users found matching your search."})
        
        user_list = []
        for user in response["value"]:
            user_summary = {
                "displayName": user.get("displayName"),
                "userPrincipalName": user.get("userPrincipalName"),
                "mail": user.get("mail"),
                "id": user.get("id")
            }
            user_list.append(user_summary)
        
        return json.dumps(user_list, indent=2)
        
    except Exception as e:
        logger.error(f"[search_users] Error: {e}")
        return f"❌ Error: {str(e)}"
