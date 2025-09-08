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
            
            # Extract attachments and file information
            attachments = resource.get("attachments", [])
            file_info = []
            
            # Process explicit attachments
            for attachment in attachments:
                attachment_info = {
                    "name": attachment.get("name"),
                    "contentType": attachment.get("contentType"),
                    "contentUrl": attachment.get("contentUrl"),
                    "content": attachment.get("content")
                }
                
                # Handle different attachment types
                if attachment.get("contentType") == "reference":
                    # File attachments (SharePoint, OneDrive files)
                    attachment_info.update({
                        "webUrl": attachment.get("content", {}).get("downloadUrl") or 
                                 attachment.get("content", {}).get("webUrl"),
                        "downloadUrl": attachment.get("content", {}).get("downloadUrl"),
                        "sharePointFileId": attachment.get("content", {}).get("uniqueId")
                    })
                elif attachment.get("contentType") == "application/vnd.microsoft.teams.file.download.info":
                    # Direct file downloads
                    attachment_info.update({
                        "downloadUrl": attachment.get("content", {}).get("downloadUrl"),
                        "uniqueId": attachment.get("content", {}).get("uniqueId")
                    })
                
                file_info.append(attachment_info)
            
            # Also extract file links from message content
            message_content = resource.get("body", {}).get("content") or ""
            if message_content:
                # Look for SharePoint/OneDrive links in the content
                import re
                
                # More comprehensive pattern for file URLs in Teams messages
                # This pattern captures the full URL including query parameters
                file_url_pattern = r'https://[^/]+\.sharepoint\.com/[^\s<>"\']*\.(xlsx?|docx?|pptx?|pdf|txt|csv|zip|rar|7z|gz|tar|msg|eml|jpg|jpeg|png|gif|bmp|tiff?|mp4|avi|mov|wmv|mp3|wav|m4a)[^\s<>"\']*'
                
                all_file_urls = re.findall(file_url_pattern, message_content, re.IGNORECASE)
                
                for full_url in all_file_urls:
                    try:
                        # Extract filename from URL (decode URL encoding)
                        import urllib.parse
                        decoded_url = urllib.parse.unquote(full_url)
                        
                        # Look for filename in different URL patterns
                        filename_patterns = [
                            r'/([^/]+\.(xlsx?|docx?|pptx?|pdf|txt|csv|zip|rar|7z|gz|tar|msg|eml|jpg|jpeg|png|gif|bmp|tiff?|mp4|avi|mov|wmv|mp3|wav|m4a))(?:[?#]|$)',
                            r'([^/\\]+\.(xlsx?|docx?|pptx?|pdf|txt|csv|zip|rar|7z|gz|tar|msg|eml|jpg|jpeg|png|gif|bmp|tiff?|mp4|avi|mov|wmv|mp3|wav|m4a))(?:[?#%]|$)'
                        ]
                        
                        filename = "Unknown File"
                        for pattern in filename_patterns:
                            filename_match = re.search(pattern, decoded_url, re.IGNORECASE)
                            if filename_match:
                                filename = filename_match.group(1)
                                break
                        
                        # Clean up filename (remove URL encoding artifacts)
                        filename = re.sub(r'%20', ' ', filename)
                        filename = re.sub(r'%[0-9A-F]{2}', '', filename)
                        
                        # Add as file info if not already present
                        existing_file = any(
                            f.get("webUrl") == full_url or 
                            f.get("downloadUrl") == full_url or
                            f.get("name") == filename 
                            for f in file_info
                        )
                        
                        if not existing_file:
                            file_info.append({
                                "name": filename,
                                "contentType": "reference",
                                "webUrl": full_url,
                                "downloadUrl": full_url,
                                "source": "message_content"
                            })
                    except Exception as url_error:
                        logger.warning(f"Error processing file URL {full_url}: {url_error}")
                        # Still add it with basic info
                        if not any(f.get("webUrl") == full_url for f in file_info):
                            file_info.append({
                                "name": "File Link",
                                "contentType": "reference", 
                                "webUrl": full_url,
                                "downloadUrl": full_url,
                                "source": "message_content"
                            })
            
            # Also check for mentions and other entities
            mentions = []
            if resource.get("mentions"):
                for mention in resource.get("mentions", []):
                    mention_info = {
                        "id": mention.get("id"),
                        "mentionText": mention.get("mentionText"),
                        "mentioned": mention.get("mentioned", {}).get("user", {}).get("displayName")
                    }
                    mentions.append(mention_info)
            
            result = {
                "id": resource.get("id"),
                "summary": hit.get("summary"),
                "rank": hit.get("rank"),
                "content": resource.get("body", {}).get("content") or "No content",
                "from": from_info.get("displayName") or "Unknown",
                "createdDateTime": resource.get("createdDateTime"),
                "chatId": resource.get("chatId"),
                "teamId": channel_identity.get("teamId"),
                "channelId": channel_identity.get("channelId"),
                "attachments": file_info,
                "hasAttachments": len(file_info) > 0,
                "mentions": mentions,
                "messageType": resource.get("messageType"),
                "webUrl": resource.get("webUrl")
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

@server.tool()
@require_teams_service("teams", "teams_read")
async def get_message_attachments(
    service,
    user_email: str,
    team_id: str,
    channel_id: str,
    message_id: str
) -> str:
    """
    Get detailed attachment information for a specific message, including download links.
    
    Args:
        user_email (str): The user's email address. Required.
        team_id (str): Team ID
        channel_id (str): Channel ID  
        message_id (str): Message ID to get attachments for
        
    Returns:
        str: JSON string containing attachment details.
    """
    logger.info(f"[get_message_attachments] Getting attachments for message {message_id}")
    
    try:
        # Get the specific message with full details
        message_response = await service.get(f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}")
        
        if not message_response:
            return json.dumps({"message": "Message not found."})
        
        # Extract attachments
        attachments = message_response.get("attachments", [])
        attachment_details = []
        
        for attachment in attachments:
            detail = {
                "id": attachment.get("id"),
                "name": attachment.get("name"),
                "contentType": attachment.get("contentType"),
                "contentUrl": attachment.get("contentUrl"),
                "content": attachment.get("content", {}),
                "attachmentType": attachment.get("contentType")
            }
            
            # Add specific URLs based on content type
            if attachment.get("contentType") == "reference":
                content = attachment.get("content", {})
                detail.update({
                    "downloadUrl": content.get("downloadUrl"),
                    "webUrl": content.get("webUrl") or content.get("downloadUrl"),
                    "sharePointFileId": content.get("uniqueId"),
                    "driveId": content.get("driveId"),
                    "itemId": content.get("itemId")
                })
            
            attachment_details.append(detail)
        
        # Also extract file links from message body
        message_content = message_response.get("body", {}).get("content") or ""
        content_files = []
        
        if message_content:
            import re
            import urllib.parse
            
            # Find file URLs in content
            file_url_pattern = r'https://[^/]+\.sharepoint\.com/[^\s<>"\']*\.(xlsx?|docx?|pptx?|pdf|txt|csv|zip|rar|7z|gz|tar|msg|eml|jpg|jpeg|png|gif|bmp|tiff?|mp4|avi|mov|wmv|mp3|wav|m4a)[^\s<>"\']*'
            
            file_urls = re.findall(file_url_pattern, message_content, re.IGNORECASE)
            
            for url in file_urls:
                try:
                    decoded_url = urllib.parse.unquote(url)
                    filename_match = re.search(r'/([^/]+\.(xlsx?|docx?|pptx?|pdf|txt|csv|zip|rar|7z|gz|tar|msg|eml|jpg|jpeg|png|gif|bmp|tiff?|mp4|avi|mov|wmv|mp3|wav|m4a))(?:[?#]|$)', decoded_url, re.IGNORECASE)
                    filename = filename_match.group(1) if filename_match else "Unknown File"
                    
                    content_files.append({
                        "name": filename,
                        "url": url,
                        "source": "message_body"
                    })
                except Exception:
                    content_files.append({
                        "name": "File Link",
                        "url": url,
                        "source": "message_body"
                    })
        
        result = {
            "messageId": message_id,
            "teamId": team_id,
            "channelId": channel_id,
            "totalAttachments": len(attachment_details),
            "totalContentFiles": len(content_files),
            "attachments": attachment_details,
            "contentFiles": content_files,
            "messageContent": message_content[:500] + "..." if len(message_content) > 500 else message_content
        }
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"[get_message_attachments] Error: {e}")
        return f"❌ Error getting message attachments: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read") 
async def search_files_in_messages(
    service,
    user_email: str,
    file_extension: str,
    limit: int = 20,
    hours: int = 168  # 1 week
) -> str:
    """
    Search specifically for files with certain extensions mentioned in Teams messages.
    
    Args:
        user_email (str): The user's email address. Required.
        file_extension (str): File extension to search for (e.g., 'xlsx', 'pdf', 'docx')
        limit (int): Maximum number of results (default: 20, max: 50)
        hours (int): Search within last N hours (default: 168 = 1 week)
        
    Returns:
        str: JSON string containing file search results with download links.
    """
    logger.info(f"[search_files_in_messages] Searching for .{file_extension} files")
    
    try:
        # Validate parameters
        if limit < 1 or limit > 50:
            limit = 20
        if hours < 1:
            hours = 168
        
        # Calculate date range
        since = datetime.now() - timedelta(hours=hours)
        since_str = since.strftime("%Y-%m-%d")
        
        # Build search query for files
        search_query = f'sent>={since_str} AND ("{file_extension}" OR hasAttachment:true)'
        
        search_request = {
            "entityTypes": ["chatMessage"],
            "query": {
                "queryString": search_query
            },
            "from": 0,
            "size": min(limit * 2, 100),  # Get more results to filter for files
            "enableTopResults": False
        }
        
        response = await service.post("/search/query", {"requests": [search_request]})
        
        if (not response.get("value") or 
            not response["value"] or 
            not response["value"][0].get("hitsContainers")):
            return json.dumps({"message": f"No messages found containing .{file_extension} files."})
        
        hits = response["value"][0]["hitsContainers"][0].get("hits", [])
        
        file_results = []
        for hit in hits:
            resource = hit.get("resource", {})
            
            # Check message content for file references
            content = resource.get("body", {}).get("content") or ""
            if file_extension.lower() not in content.lower():
                continue
            
            # Extract file information
            import re
            import urllib.parse
            
            # Find URLs with the specific extension
            extension_pattern = rf'https://[^/]+\.sharepoint\.com/[^\s<>"\']*\.{re.escape(file_extension)}[^\s<>"\']*'
            file_urls = re.findall(extension_pattern, content, re.IGNORECASE)
            
            for url in file_urls:
                try:
                    decoded_url = urllib.parse.unquote(url)
                    filename_pattern = rf'/([^/]+\.{re.escape(file_extension)})(?:[?#]|$)'
                    filename_match = re.search(filename_pattern, decoded_url, re.IGNORECASE)
                    filename = filename_match.group(1) if filename_match else f"file.{file_extension}"
                    
                    from_info = resource.get("from", {}).get("user", {})
                    
                    file_results.append({
                        "filename": filename,
                        "downloadUrl": url,
                        "messageId": resource.get("id"),
                        "from": from_info.get("displayName") or "Unknown",
                        "createdDateTime": resource.get("createdDateTime"),
                        "teamId": resource.get("channelIdentity", {}).get("teamId"),
                        "channelId": resource.get("channelIdentity", {}).get("channelId"),
                        "chatId": resource.get("chatId")
                    })
                except Exception:
                    pass
            
            # Stop if we have enough results
            if len(file_results) >= limit:
                break
        
        result = {
            "fileExtension": file_extension,
            "timeRange": f"Last {hours} hours", 
            "totalFilesFound": len(file_results[:limit]),
            "files": file_results[:limit]
        }
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"[search_files_in_messages] Error: {e}")
        return f"❌ Error searching for files: {str(e)}"
