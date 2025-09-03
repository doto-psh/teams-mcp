"""
Microsoft Teams OAuth Scopes

This module centralizes OAuth scope definitions for Microsoft Teams integration.
Adapted from Google Workspace MCP for Microsoft Graph API.
"""
import logging

logger = logging.getLogger(__name__)

# Global variable to store enabled tools (set by main.py)
_ENABLED_TOOLS = None

# Base Microsoft Graph API scopes
# Note: Microsoft Graph uses its own scopes, not OIDC standard scopes
USER_READ_SCOPE = 'https://graph.microsoft.com/User.Read'

# Microsoft Teams scopes
TEAMS_READ_SCOPE = 'https://graph.microsoft.com/Team.ReadBasic.All'
TEAMS_CHANNELS_READ_SCOPE = 'https://graph.microsoft.com/Channel.ReadBasic.All'
TEAMS_MESSAGES_READ_SCOPE = 'https://graph.microsoft.com/ChannelMessage.Read.All'
TEAMS_CHAT_READ_SCOPE = 'https://graph.microsoft.com/Chat.Read'
TEAMS_CHAT_READWRITE_SCOPE = 'https://graph.microsoft.com/Chat.ReadWrite'
TEAMS_MEMBERS_READ_SCOPE = 'https://graph.microsoft.com/TeamMember.Read.All'

# Additional Microsoft Graph scopes for user management
USER_READ_SCOPE = 'https://graph.microsoft.com/User.Read'
USER_READ_ALL_SCOPE = 'https://graph.microsoft.com/User.Read.All'
DIRECTORY_READ_SCOPE = 'https://graph.microsoft.com/Directory.Read.All'

# Slides API scopes (not applicable for Teams)
# SLIDES_SCOPE = 'https://www.googleapis.com/auth/presentations'
# SLIDES_READONLY_SCOPE = 'https://www.googleapis.com/auth/presentations.readonly'

# Tasks API scopes (not applicable for Teams)
# TASKS_SCOPE = 'https://www.googleapis.com/auth/tasks'
# TASKS_READONLY_SCOPE = 'https://www.googleapis.com/auth/tasks.readonly'

# Custom Search API scope (not applicable for Teams)
# CUSTOM_SEARCH_SCOPE = 'https://www.googleapis.com/auth/cse'

# Base OAuth scopes required for user identification
# Note: Microsoft Graph API doesn't support mixing OIDC scopes with Graph scopes
BASE_SCOPES = [
    USER_READ_SCOPE,  # This provides user profile information
]

# Service-specific scope groups for Teams
TEAMS_SCOPES = [
    TEAMS_READ_SCOPE,
    TEAMS_CHANNELS_READ_SCOPE,
    TEAMS_MESSAGES_READ_SCOPE,
    TEAMS_CHAT_READ_SCOPE,
    TEAMS_MEMBERS_READ_SCOPE,
    USER_READ_SCOPE
]

# Tool-to-scopes mapping for Teams MCP
TOOL_SCOPES_MAP = {
    'teams': TEAMS_SCOPES,
    'user': [USER_READ_SCOPE, USER_READ_ALL_SCOPE]
}

def set_enabled_tools(enabled_tools):
    """
    Set the globally enabled tools list.
    
    Args:
        enabled_tools: List of enabled tool names.
    """
    global _ENABLED_TOOLS
    _ENABLED_TOOLS = enabled_tools
    logger.info(f"Enabled tools set for scope management: {enabled_tools}")

def get_current_scopes():
    """
    Returns scopes for currently enabled tools.
    Uses globally set enabled tools or all tools if not set.
    
    Returns:
        List of unique scopes for the enabled tools plus base scopes.
    """
    enabled_tools = _ENABLED_TOOLS
    if enabled_tools is None:
        # Default behavior - return Teams scopes only
        enabled_tools = ['teams']
    
    # Start with base scopes (always required)
    scopes = BASE_SCOPES.copy()
    
    # Add scopes for each enabled tool
    for tool in enabled_tools:
        if tool in TOOL_SCOPES_MAP:
            scopes.extend(TOOL_SCOPES_MAP[tool])
    
    logger.debug(f"Generated scopes for tools {list(enabled_tools)}: {len(set(scopes))} unique scopes")
    # Return unique scopes
    return list(set(scopes))

def get_scopes_for_tools(enabled_tools=None):
    """
    Returns scopes for enabled tools only.
    
    Args:
        enabled_tools: List of enabled tool names. If None, returns Teams scopes.
    
    Returns:
        List of unique scopes for the enabled tools plus base scopes.
    """
    if enabled_tools is None:
        # Default behavior - return Teams scopes only
        enabled_tools = ['teams']
    
    # Start with base scopes (always required)
    scopes = BASE_SCOPES.copy()
    
    # Add scopes for each enabled tool
    for tool in enabled_tools:
        if tool in TOOL_SCOPES_MAP:
            scopes.extend(TOOL_SCOPES_MAP[tool])
    
    # Return unique scopes
    return list(set(scopes))

# Combined scopes for Microsoft Teams operations (backwards compatibility)
SCOPES = get_scopes_for_tools(['teams'])