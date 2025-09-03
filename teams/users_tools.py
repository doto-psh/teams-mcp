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

@server.tool()
@require_teams_service("teams", "teams_read")
async def get_user(service, user_email: str, user_id: str) -> str:
    """
    Get detailed information about a specific user by their ID or email address. Returns profile information including name, email, job title, and department.
    
    Args:
        user_email (str): The user's email address. Required.
        user_id (str): User ID or email address
        
    Returns:
        str: JSON string containing user profile information.
    """
    logger.info(f"[get_user] Fetching user info for {user_id}, user: {user_email}")
    
    try:
        user = await service.get(f"/users/{user_id}")
        
        user_summary = {
            "displayName": user.get("displayName"),
            "userPrincipalName": user.get("userPrincipalName"),
            "mail": user.get("mail"),
            "id": user.get("id"),
            "jobTitle": user.get("jobTitle"),
            "department": user.get("department"),
            "officeLocation": user.get("officeLocation")
        }
        
        return json.dumps(user_summary, indent=2)
        
    except Exception as e:
        logger.error(f"[get_user] Error: {e}")
        return f"❌ Error: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def get_user_manager(service, user_email: str, user_id: Optional[str] = None) -> str:
    """
    Get the manager of a specific user or the current user. Returns manager's profile information.
    
    Args:
        user_email (str): The user's email address. Required.
        user_id (str): User ID or email address (optional, defaults to current user)
        
    Returns:
        str: JSON string containing manager information.
    """
    logger.info(f"[get_user_manager] Fetching manager for user {user_id or 'current user'}, user: {user_email}")
    
    try:
        # Use current user if no user_id specified
        endpoint = f"/users/{user_id}/manager" if user_id else "/me/manager"
        
        manager = await service.get(endpoint)
        
        manager_summary = {
            "displayName": manager.get("displayName"),
            "userPrincipalName": manager.get("userPrincipalName"),
            "mail": manager.get("mail"),
            "id": manager.get("id"),
            "jobTitle": manager.get("jobTitle"),
            "department": manager.get("department")
        }
        
        return json.dumps(manager_summary, indent=2)
        
    except Exception as e:
        logger.error(f"[get_user_manager] Error: {e}")
        if "does not exist" in str(e).lower() or "not found" in str(e).lower():
            return json.dumps({"message": "No manager found for this user."})
        return f"❌ Error: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def get_user_direct_reports(service, user_email: str, user_id: Optional[str] = None) -> str:
    """
    Get the direct reports of a specific user or the current user. Returns list of direct reports with their profile information.
    
    Args:
        user_email (str): The user's email address. Required.
        user_id (str): User ID or email address (optional, defaults to current user)
        
    Returns:
        str: JSON string containing direct reports information.
    """
    logger.info(f"[get_user_direct_reports] Fetching direct reports for user {user_id or 'current user'}, user: {user_email}")
    
    try:
        # Use current user if no user_id specified
        endpoint = f"/users/{user_id}/directReports" if user_id else "/me/directReports"
        
        response = await service.get(endpoint)
        
        if not response.get("value"):
            return json.dumps({"message": "No direct reports found for this user."})
        
        direct_reports = []
        for report in response["value"]:
            report_summary = {
                "displayName": report.get("displayName"),
                "userPrincipalName": report.get("userPrincipalName"),
                "mail": report.get("mail"),
                "id": report.get("id"),
                "jobTitle": report.get("jobTitle"),
                "department": report.get("department")
            }
            direct_reports.append(report_summary)
        
        result = {
            "totalDirectReports": len(direct_reports),
            "directReports": direct_reports
        }
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"[get_user_direct_reports] Error: {e}")
        return f"❌ Error: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def get_user_photo(service, user_email: str, user_id: Optional[str] = None) -> str:
    """
    Get the profile photo information for a specific user or the current user.
    
    Args:
        user_email (str): The user's email address. Required.
        user_id (str): User ID or email address (optional, defaults to current user)
        
    Returns:
        str: JSON string containing photo information or message if no photo exists.
    """
    logger.info(f"[get_user_photo] Fetching photo info for user {user_id or 'current user'}, user: {user_email}")
    
    try:
        # Use current user if no user_id specified
        endpoint = f"/users/{user_id}/photo" if user_id else "/me/photo"
        
        photo_info = await service.get(endpoint)
        
        photo_summary = {
            "id": photo_info.get("id"),
            "width": photo_info.get("width"),
            "height": photo_info.get("height"),
            "type": photo_info.get("@odata.mediaContentType"),
            "hasPhoto": True
        }
        
        return json.dumps(photo_summary, indent=2)
        
    except Exception as e:
        logger.error(f"[get_user_photo] Error: {e}")
        if "does not exist" in str(e).lower() or "not found" in str(e).lower():
            return json.dumps({"hasPhoto": False, "message": "No profile photo found for this user."})
        return f"❌ Error: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def get_organization_users(service, user_email: str, limit: int = 50) -> str:
    """
    Get a list of users in the organization with basic profile information.
    
    Args:
        user_email (str): The user's email address. Required.
        limit (int): Maximum number of users to return (default: 50, max: 100)
        
    Returns:
        str: JSON string containing organization users.
    """
    logger.info(f"[get_organization_users] Fetching organization users, limit: {limit}, user: {user_email}")
    
    try:
        # Validate limit
        if limit < 1 or limit > 100:
            limit = 50
        
        response = await service.get(f"/users?$top={limit}&$select=displayName,userPrincipalName,mail,id,jobTitle,department")
        
        if not response.get("value"):
            return json.dumps({"message": "No users found in the organization."})
        
        user_list = []
        for user in response["value"]:
            user_summary = {
                "displayName": user.get("displayName"),
                "userPrincipalName": user.get("userPrincipalName"),
                "mail": user.get("mail"),
                "id": user.get("id"),
                "jobTitle": user.get("jobTitle"),
                "department": user.get("department")
            }
            user_list.append(user_summary)
        
        result = {
            "totalReturned": len(user_list),
            "hasMore": bool(response.get("@odata.nextLink")),
            "users": user_list
        }
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"[get_organization_users] Error: {e}")
        return f"❌ Error: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def search_users_advanced(service, user_email: str, display_name: Optional[str] = None, 
                               department: Optional[str] = None, job_title: Optional[str] = None, 
                               limit: int = 25) -> str:
    """
    Advanced search for users in the organization using multiple criteria.
    
    Args:
        user_email (str): The user's email address. Required.
        display_name (str): Filter by display name (contains search)
        department (str): Filter by department
        job_title (str): Filter by job title (contains search)
        limit (int): Maximum number of users to return (default: 25, max: 100)
        
    Returns:
        str: JSON string containing matching users.
    """
    logger.info(f"[search_users_advanced] Advanced user search, user: {user_email}")
    
    try:
        # Validate limit
        if limit < 1 or limit > 100:
            limit = 25
        
        # Build filter conditions
        filter_conditions = []
        
        if display_name:
            filter_conditions.append(f"contains(displayName,'{display_name}')")
        
        if department:
            filter_conditions.append(f"department eq '{department}'")
        
        if job_title:
            filter_conditions.append(f"contains(jobTitle,'{job_title}')")
        
        # Build query
        query_params = [f"$top={limit}"]
        
        if filter_conditions:
            filter_query = " and ".join(filter_conditions)
            query_params.append(f"$filter={filter_query}")
        
        query_params.append("$select=displayName,userPrincipalName,mail,id,jobTitle,department,officeLocation")
        
        query_string = "&".join(query_params)
        response = await service.get(f"/users?{query_string}")
        
        if not response.get("value"):
            return json.dumps({"message": "No users found matching the search criteria."})
        
        user_list = []
        for user in response["value"]:
            user_summary = {
                "displayName": user.get("displayName"),
                "userPrincipalName": user.get("userPrincipalName"),
                "mail": user.get("mail"),
                "id": user.get("id"),
                "jobTitle": user.get("jobTitle"),
                "department": user.get("department"),
                "officeLocation": user.get("officeLocation")
            }
            user_list.append(user_summary)
        
        result = {
            "searchCriteria": {
                "displayName": display_name,
                "department": department,
                "jobTitle": job_title
            },
            "totalFound": len(user_list),
            "hasMore": bool(response.get("@odata.nextLink")),
            "users": user_list
        }
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"[search_users_advanced] Error: {e}")
        return f"❌ Error: {str(e)}"
