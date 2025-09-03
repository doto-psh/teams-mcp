"""
Microsoft Teams Search Tools

This module provides MCP tools for searching messages across Microsoft Teams via Graph API.
"""

import json
import logging
from typing import List, Dict, Any, Optional
from datetime import datetime, timedelta
from auth.service_decorator_teams import require_teams_service
from core.server import server

logger = logging.getLogger(__name__)

@server.tool()
@require_teams_service("teams", "teams_read")
async def search_messages(
    service,
    user_email: str,
    query: str,
    scope: str = "all",
    limit: int = 25,
    enable_top_results: bool = True
) -> str:
    """
    Search for messages across all Microsoft Teams channels and chats using Microsoft Search API. Supports advanced KQL syntax for filtering by sender, mentions, attachments, and more.
    
    Args:
        user_email (str): The user's email address. Required.
        query (str): Search query. Supports KQL syntax like 'from:user mentions:userId hasAttachment:true'
        scope (str): Scope of search (all, channels, chats)
        limit (int): Number of results to return (default: 25, max: 100)
        enable_top_results (bool): Enable relevance-based ranking
        
    Returns:
        str: JSON string containing search results.
    """
    logger.info(f"[search_messages] Searching messages with query '{query}', user: {user_email}")
    
    try:
        # Validate limit
        if limit < 1 or limit > 100:
            limit = 25
        
        # Build the search request
        search_request = {
            "entityTypes": ["chatMessage"],
            "query": {
                "queryString": query
            },
            "from": 0,
            "size": limit,
            "enableTopResults": enable_top_results
        }
        
        # Add scope-specific filters to the query if needed
        enhanced_query = query
        if scope == "channels":
            enhanced_query = f"{query} AND (channelIdentity/channelId:*)"
        elif scope == "chats":
            enhanced_query = f"{query} AND (chatId:* AND NOT channelIdentity/channelId:*)"
        
        search_request["query"]["queryString"] = enhanced_query
        
        response = await service.post("/search/query", {"requests": [search_request]})
        
        if (not response.get("value") or 
            not response["value"] or 
            not response["value"][0].get("hitsContainers")):
            return json.dumps({"message": "No messages found matching your search criteria."})
        
        hits = response["value"][0]["hitsContainers"][0].get("hits", [])
        
        search_results = []
        for hit in hits:
            resource = hit.get("resource", {})
            channel_identity = resource.get("channelIdentity", {})
            from_info = resource.get("from", {}).get("user", {})
            
            result = {
                "id": resource.get("id"),
                "summary": hit.get("summary"),
                "rank": hit.get("rank"),
                "content": resource.get("body", {}).get("content") or "No content",
                "from": from_info.get("displayName") or "Unknown",
                "createdDateTime": resource.get("createdDateTime"),
                "chatId": resource.get("chatId"),
                "teamId": channel_identity.get("teamId"),
                "channelId": channel_identity.get("channelId")
            }
            search_results.append(result)
        
        result = {
            "query": query,
            "scope": scope,
            "totalResults": response["value"][0]["hitsContainers"][0].get("total", 0),
            "results": search_results,
            "moreResultsAvailable": response["value"][0]["hitsContainers"][0].get("moreResultsAvailable", False)
        }
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"[search_messages] Error: {e}")
        return f"❌ Error searching messages: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def get_recent_messages(
    service,
    user_email: str,
    hours: int = 24,
    limit: int = 50,
    mentions_user: Optional[str] = None,
    from_user: Optional[str] = None,
    has_attachments: Optional[bool] = None,
    importance: Optional[str] = None,
    include_channels: bool = True,
    include_chats: bool = True,
    team_ids: Optional[List[str]] = None,
    keywords: Optional[str] = None
) -> str:
    """
    Get recent messages from across Teams with advanced filtering options. Can filter by time range, scope (channels vs chats), teams, channels, and users.
    
    Args:
        user_email (str): The user's email address. Required.
        hours (int): Get messages from the last N hours (max 168 = 1 week)
        limit (int): Maximum number of messages to return (default: 50, max: 100)
        mentions_user (str): Filter messages that mention this user ID
        from_user (str): Filter messages from this user ID
        has_attachments (bool): Filter messages with attachments
        importance (str): Filter by message importance (low, normal, high, urgent)
        include_channels (bool): Include channel messages
        include_chats (bool): Include chat messages
        team_ids (List[str]): Specific team IDs to search in
        keywords (str): Keywords to search for in message content
        
    Returns:
        str: JSON string containing recent messages.
    """
    logger.info(f"[get_recent_messages] Fetching recent messages for last {hours} hours, user: {user_email}")
    
    try:
        # Validate parameters
        if hours < 1 or hours > 168:
            hours = 24
        if limit < 1 or limit > 100:
            limit = 50
        
        attempted_advanced_search = False
        
        # Try using the Search API first for rich filtering
        if keywords or mentions_user or has_attachments is not None or importance:
            attempted_advanced_search = True
            
            # Calculate the date threshold
            since = datetime.now() - timedelta(hours=hours)
            since_str = since.strftime("%Y-%m-%d")
            
            # Build KQL query for Microsoft Search API
            query_parts = [f"sent>={since_str}"]  # Use just the date part
            
            # Add user filters
            if mentions_user:
                query_parts.append(f"mentions:{mentions_user}")
            if from_user:
                query_parts.append(f"from:{from_user}")
            
            # Add content filters
            if has_attachments is not None:
                query_parts.append(f"hasAttachment:{str(has_attachments).lower()}")
            if importance:
                query_parts.append(f"importance:{importance}")
            
            # Add keyword search
            if keywords:
                query_parts.append(f'"{keywords}"')
            
            # If no specific filters, search for all recent messages
            if len(query_parts) == 1:
                query_parts.append("*")  # Match all messages
            
            search_query = " AND ".join(query_parts)
            
            search_request = {
                "entityTypes": ["chatMessage"],
                "query": {
                    "queryString": search_query
                },
                "from": 0,
                "size": min(limit, 100),
                "enableTopResults": False  # For recent messages, prefer chronological order
            }
            
            try:
                response = await service.post("/search/query", {"requests": [search_request]})
                
                if (response.get("value") and 
                    response["value"] and 
                    response["value"][0].get("hitsContainers")):
                    
                    hits = response["value"][0]["hitsContainers"][0].get("hits", [])
                    
                    # Filter and process results
                    recent_messages = []
                    for hit in hits:
                        resource = hit.get("resource", {})
                        channel_identity = resource.get("channelIdentity", {})
                        
                        # Apply scope filters
                        is_channel_message = bool(channel_identity.get("channelId"))
                        is_chat_message = bool(resource.get("chatId") and not is_channel_message)
                        
                        if not include_channels and is_channel_message:
                            continue
                        if not include_chats and is_chat_message:
                            continue
                        
                        # Apply team filter if specified
                        if team_ids and is_channel_message:
                            if channel_identity.get("teamId") not in team_ids:
                                continue
                        
                        from_info = resource.get("from", {}).get("user", {})
                        message = {
                            "id": resource.get("id"),
                            "content": resource.get("body", {}).get("content") or "No content",
                            "from": from_info.get("displayName") or "Unknown",
                            "fromUserId": from_info.get("id"),
                            "createdDateTime": resource.get("createdDateTime"),
                            "chatId": resource.get("chatId"),
                            "teamId": channel_identity.get("teamId"),
                            "channelId": channel_identity.get("channelId"),
                            "type": "channel" if is_channel_message else "chat"
                        }
                        recent_messages.append(message)
                    
                    # Apply final limit after filtering
                    recent_messages = recent_messages[:limit]
                    
                    # Check if Search API returned poor quality results
                    poor_quality_results = sum(1 for msg in recent_messages 
                                             if msg["content"] == "No content" or msg["from"] == "Unknown")
                    
                    quality_threshold = 0.5  # If more than 50% are poor quality, fall back
                    if (recent_messages and 
                        poor_quality_results / len(recent_messages) <= quality_threshold):
                        
                        result = {
                            "method": "search_api",
                            "timeRange": f"Last {hours} hours",
                            "filters": {
                                "mentionsUser": mentions_user,
                                "fromUser": from_user,
                                "hasAttachments": has_attachments,
                                "importance": importance,
                                "keywords": keywords
                            },
                            "totalFound": len(recent_messages),
                            "messages": recent_messages
                        }
                        
                        return json.dumps(result, indent=2)
                        
            except Exception as search_error:
                logger.error(f"Search API failed, falling back to direct queries: {search_error}")
        
        # Fallback: Get recent messages from user's chats directly
        chats_response = await service.get("/me/chats?$expand=members")
        chats = chats_response.get("value", [])
        
        all_messages = []
        since = datetime.now() - timedelta(hours=hours)
        
        # Get recent messages from each chat (limit to first 10 chats to avoid rate limits)
        for chat in chats[:10]:
            try:
                query_string = f"$top={min(limit, 50)}&$orderby=createdDateTime desc"
                
                # Apply user filter if specified
                if from_user:
                    query_string += f"&$filter=from/user/id eq '{from_user}'"
                
                messages_response = await service.get(f"/me/chats/{chat['id']}/messages?{query_string}")
                messages = messages_response.get("value", [])
                
                for message in messages:
                    # Filter by time
                    if message.get("createdDateTime"):
                        try:
                            message_date = datetime.fromisoformat(message["createdDateTime"].replace('Z', '+00:00'))
                            if message_date < since:
                                continue
                        except (ValueError, AttributeError):
                            continue
                    
                    # Apply scope filter for chats
                    if not include_chats:
                        continue
                    
                    # Apply keyword filter (simple text search)
                    if keywords and message.get("body", {}).get("content"):
                        content = message["body"]["content"].lower()
                        if keywords.lower() not in content:
                            continue
                    
                    from_info = message.get("from", {}).get("user", {})
                    all_messages.append({
                        "id": message.get("id", ""),
                        "content": message.get("body", {}).get("content") or "No content",
                        "from": from_info.get("displayName") or "Unknown",
                        "fromUserId": from_info.get("id"),
                        "createdDateTime": message.get("createdDateTime", ""),
                        "chatId": message.get("chatId", ""),
                        "type": "chat"
                    })
                    
                    if len(all_messages) >= limit:
                        break
                
                if len(all_messages) >= limit:
                    break
                    
            except Exception as chat_error:
                logger.error(f"Error getting messages from chat {chat.get('id')}: {chat_error}")
        
        # Sort by creation date (newest first)
        all_messages.sort(key=lambda x: x.get("createdDateTime", ""), reverse=True)
        
        result = {
            "method": "direct_chat_queries_fallback" if attempted_advanced_search else "direct_chat_queries",
            "timeRange": f"Last {hours} hours",
            "filters": {
                "mentionsUser": mentions_user,
                "fromUser": from_user,
                "hasAttachments": has_attachments,
                "importance": importance,
                "keywords": keywords
            },
            "note": ("Search API returned poor quality results, using direct chat queries as fallback" 
                    if attempted_advanced_search 
                    else "Using direct chat queries for better content reliability"),
            "totalFound": len(all_messages[:limit]),
            "messages": all_messages[:limit]
        }
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"[get_recent_messages] Error: {e}")
        return f"❌ Error getting recent messages: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def get_my_mentions(
    service,
    user_email: str,
    hours: int = 24,
    limit: int = 20,
    scope: str = "all"
) -> str:
    """
    Find all recent messages where the current user was mentioned (@mentioned) across Teams channels and chats.
    
    Args:
        user_email (str): The user's email address. Required.
        hours (int): Get mentions from the last N hours (default: 24, max: 168)
        limit (int): Maximum number of mentions to return (default: 20, max: 50)
        scope (str): Scope of search (all, channels, chats)
        
    Returns:
        str: JSON string containing mention results.
    """
    logger.info(f"[get_my_mentions] Fetching mentions for last {hours} hours, user: {user_email}")
    
    try:
        # Validate parameters
        if hours < 1 or hours > 168:
            hours = 24
        if limit < 1 or limit > 50:
            limit = 20
        
        # Get current user ID first
        me = await service.get("/me")
        user_id = me.get("id")
        
        if not user_id:
            return "❌ Error: Could not determine current user ID"
        
        since = datetime.now() - timedelta(hours=hours)
        since_str = since.strftime("%Y-%m-%d")  # Use just the date part to avoid time parsing issues
        
        # Build query to find mentions of current user
        query_parts = [
            f"sent>={since_str}",
            f"mentions:{user_id}"
        ]
        
        search_query = " AND ".join(query_parts)
        
        search_request = {
            "entityTypes": ["chatMessage"],
            "query": {
                "queryString": search_query
            },
            "from": 0,
            "size": min(limit, 50),
            "enableTopResults": False
        }
        
        response = await service.post("/search/query", {"requests": [search_request]})
        
        if (not response.get("value") or 
            not response["value"] or 
            not response["value"][0].get("hitsContainers") or
            not response["value"][0]["hitsContainers"][0].get("hits")):
            return json.dumps({"message": "No recent mentions found."})
        
        hits = response["value"][0]["hitsContainers"][0]["hits"]
        
        if not hits:
            return json.dumps({"message": "No recent mentions found."})
        
        mentions = []
        for hit in hits:
            resource = hit.get("resource", {})
            channel_identity = resource.get("channelIdentity", {})
            
            # Apply scope filters
            is_channel_message = bool(channel_identity.get("channelId"))
            is_chat_message = bool(resource.get("chatId") and not is_channel_message)
            
            if scope == "channels" and not is_channel_message:
                continue
            if scope == "chats" and not is_chat_message:
                continue
            
            from_info = resource.get("from", {}).get("user", {})
            mention = {
                "id": resource.get("id"),
                "content": resource.get("body", {}).get("content") or "No content",
                "summary": hit.get("summary"),
                "from": from_info.get("displayName") or "Unknown",
                "fromUserId": from_info.get("id"),
                "createdDateTime": resource.get("createdDateTime"),
                "chatId": resource.get("chatId"),
                "teamId": channel_identity.get("teamId"),
                "channelId": channel_identity.get("channelId"),
                "type": "channel" if is_channel_message else "chat"
            }
            mentions.append(mention)
        
        result = {
            "timeRange": f"Last {hours} hours",
            "mentionedUser": me.get("displayName") or "Current User",
            "scope": scope,
            "totalMentions": len(mentions),
            "mentions": mentions
        }
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"[get_my_mentions] Error: {e}")
        return f"❌ Error getting mentions: {str(e)}"
