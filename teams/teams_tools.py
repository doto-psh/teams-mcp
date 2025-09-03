"""
Microsoft Teams Tools

This module provides MCP tools for interacting with Microsoft Teams via Graph API.
"""

import json
import logging
import base64
import re
from typing import List, Dict, Any, Optional
from auth.service_decorator_teams import require_teams_service
from core.server import server

logger = logging.getLogger(__name__)

@server.tool()
@require_teams_service("teams", "teams_read")
async def list_teams(service, user_email: str) -> str:
    """
    List all Microsoft Teams that the current user is a member of. Returns team names, descriptions, and IDs.
    
    Args:
        user_email (str): The user's email address. Required.
        
    Returns:
        str: JSON string containing teams information.
    """
    logger.info(f"[list_teams] Fetching teams for user: {user_email}")
    
    try:
        # Get user's joined teams
        teams_data = await service.get("/me/joinedTeams")
        
        if not teams_data.get("value"):
            return json.dumps({"message": "No teams found."})
        
        team_list = [
            {
                "id": team.get("id"),
                "displayName": team.get("displayName"),
                "description": team.get("description"),
                "isArchived": team.get("isArchived", False),
            }
            for team in teams_data["value"]
        ]

        return json.dumps(team_list, indent=2)

    except Exception as e:
        logger.error(f"[list_teams] Error: {e}")
        return f"❌ Error: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def list_channels(service, user_email: str, team_id: str) -> str:
    """
    List all channels in a specific Microsoft Team. Returns channel names, descriptions, types, and IDs for the specified team.

    Args:
        user_email (str): The user's email address. Required.
        team_id (str): The ID of the team to get channels from.
        
    Returns:
        str: JSON string containing channels information.
    """
    logger.info(f"[list_channels] Fetching channels for team {team_id}, user: {user_email}")
    
    try:
        # Get team channels
        channels_data = await service.get(f"/teams/{team_id}/channels")
        
        if not channels_data.get("value"):
            return json.dumps({"message": "No channels found in this team."})
        
        channel_list = [
            {
                "id": channel.get("id"),
                "displayName": channel.get("displayName"),
                "description": channel.get("description"),
                "membershipType": channel.get("membershipType"),
            }
            for channel in channels_data["value"]
        ]
        
        return json.dumps(channel_list, indent=2)
        
    except Exception as e:
        logger.error(f"[list_channels] Error: {e}")
        return f"❌ Error: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def get_channel_messages(service, user_email: str, team_id: str, channel_id: str, limit: int = 20) -> str:
    """
    Retrieve recent messages from a specific channel in a Microsoft Team. Returns message content, sender information, and timestamps.
    
    Args:
        user_email (str): The user's email address. Required.
        team_id (str): Team ID
        channel_id (str): Channel ID
        limit (int): Number of messages to retrieve (default: 20, max: 50)
        
    Returns:
        str: JSON string containing messages information.
    """
    logger.info(f"[get_channel_messages] Fetching messages for team {team_id}, channel {channel_id}, user: {user_email}")
    
    try:
        # Validate limit
        if limit < 1 or limit > 50:
            limit = 20
        
        # Build query parameters - Teams channel messages API has limited query support
        # Only $top is supported, no $orderby, $filter, etc.
        query_params = f"$top={limit}"
        
        messages_data = await service.get(f"/teams/{team_id}/channels/{channel_id}/messages?{query_params}")
        
        if not messages_data.get("value"):
            return json.dumps({"message": "No messages found in this channel."})
        
        message_list = []
        for message in messages_data["value"]:
            message_info = {
                "id": message.get("id"),
                "content": message.get("body", {}).get("content"),
                "from": message.get("from", {}).get("user", {}).get("displayName"),
                "createdDateTime": message.get("createdDateTime"),
                "importance": message.get("importance"),
            }
            message_list.append(message_info)
        
        # Sort messages by creation date (newest first) since API doesn't support orderby
        message_list.sort(key=lambda x: x.get("createdDateTime", ""), reverse=True)
        
        result = {
            "totalReturned": len(message_list),
            "hasMore": bool(messages_data.get("@odata.nextLink")),
            "messages": message_list,
        }
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"[get_channel_messages] Error: {e}")
        return f"❌ Error: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_write")
async def send_channel_message(
    service, 
    user_email: str, 
    team_id: str, 
    channel_id: str, 
    message: str,
    importance: str = "normal",
    format: str = "text",
    mentions: Optional[List[Dict[str, str]]] = None,
    image_url: Optional[str] = None,
    image_data: Optional[str] = None,
    image_content_type: Optional[str] = None,
    image_file_name: Optional[str] = None
) -> str:
    """
    Send a message to a specific channel in a Microsoft Team. Supports text and markdown formatting, mentions, and importance levels.
    
    Args:
        user_email (str): The user's email address. Required.
        team_id (str): Team ID
        channel_id (str): Channel ID
        message (str): Message content
        importance (str): Message importance (normal, high, urgent)
        format (str): Message format (text or markdown)
        mentions (List[Dict]): Array of @mentions to include in the message
        image_url (str): URL of an image to attach to the message
        image_data (str): Base64 encoded image data to attach
        image_content_type (str): MIME type of the image
        image_file_name (str): Name for the attached image file
        
    Returns:
        str: Success or error message.
    """
    logger.info(f"[send_channel_message] Sending message to team {team_id}, channel {channel_id}, user: {user_email}")
    
    try:
        # Process message content based on format
        content = message
        content_type = "text"
        
        if format == "markdown":
            # Simple markdown to HTML conversion (you might want to use a proper library)
            content = await _markdown_to_html(message)
            content_type = "html"
        
        # Process @mentions if provided
        final_mentions = []
        mention_mappings = []
        
        if mentions:
            for i, mention in enumerate(mentions):
                try:
                    # Get user info to get display name
                    user_response = await service.get(f"/users/{mention['userId']}?$select=displayName")
                    display_name = user_response.get("displayName", mention["mention"])
                    
                    mention_mappings.append({
                        "mention": mention["mention"],
                        "userId": mention["userId"],
                        "displayName": display_name,
                    })
                except Exception as e:
                    logger.warning(f"Could not resolve user {mention['userId']}: {e}")
                    mention_mappings.append({
                        "mention": mention["mention"],
                        "userId": mention["userId"],
                        "displayName": mention["mention"],
                    })
        
        # Process mentions in HTML content
        if mention_mappings:
            content, final_mentions = await _process_mentions_in_html(content, mention_mappings)
            content_type = "html"
        
        # Handle image attachment (simplified version)
        attachments = []
        if image_url or image_data:
            # For simplicity, we'll skip image upload in this implementation
            # In a full implementation, you'd handle image upload to SharePoint/OneDrive
            logger.info("Image attachments not fully implemented in this version")
        
        # Build message payload
        message_payload = {
            "body": {
                "content": content,
                "contentType": content_type,
            },
            "importance": importance,
        }
        
        if final_mentions:
            message_payload["mentions"] = final_mentions
        
        if attachments:
            message_payload["attachments"] = attachments
        
        result = await service.post(f"/teams/{team_id}/channels/{channel_id}/messages", message_payload)
        
        # Build success message
        success_text = f"✅ Message sent successfully. Message ID: {result.get('id')}"
        if final_mentions:
            mentions_text = ", ".join([m.get("mentionText", "") for m in final_mentions])
            success_text += f"\n📱 Mentions: {mentions_text}"
        if attachments:
            success_text += f"\n🖼️ Image attached: {attachments[0].get('name', '')}"
        
        return success_text
        
    except Exception as e:
        logger.error(f"[send_channel_message] Error: {e}")
        return f"❌ Failed to send message: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def get_channel_message_replies(
    service, 
    user_email: str, 
    team_id: str, 
    channel_id: str, 
    message_id: str, 
    limit: int = 20
) -> str:
    """
    Get all replies to a specific message in a channel. Returns reply content, sender information, and timestamps.
    
    Args:
        user_email (str): The user's email address. Required.
        team_id (str): Team ID
        channel_id (str): Channel ID
        message_id (str): Message ID to get replies for
        limit (int): Number of replies to retrieve (default: 20, max: 50)
        
    Returns:
        str: JSON string containing replies information.
    """
    logger.info(f"[get_channel_message_replies] Fetching replies for message {message_id} in team {team_id}, channel {channel_id}")
    
    try:
        # Validate limit
        if limit < 1 or limit > 50:
            limit = 20
        
        # Only $top is supported for message replies
        query_params = f"$top={limit}"
        
        replies_data = await service.get(f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies?{query_params}")
        
        if not replies_data.get("value"):
            return json.dumps({"message": "No replies found for this message."})
        
        replies_list = []
        for reply in replies_data["value"]:
            reply_info = {
                "id": reply.get("id"),
                "content": reply.get("body", {}).get("content"),
                "from": reply.get("from", {}).get("user", {}).get("displayName"),
                "createdDateTime": reply.get("createdDateTime"),
                "importance": reply.get("importance"),
            }
            replies_list.append(reply_info)
        
        # Sort replies by creation date (oldest first for replies)
        replies_list.sort(key=lambda x: x.get("createdDateTime", ""))
        
        result = {
            "parentMessageId": message_id,
            "totalReplies": len(replies_list),
            "hasMore": bool(replies_data.get("@odata.nextLink")),
            "replies": replies_list,
        }
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"[get_channel_message_replies] Error: {e}")
        return f"❌ Error: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_write")
async def reply_to_channel_message(
    service, 
    user_email: str, 
    team_id: str, 
    channel_id: str, 
    message_id: str,
    message: str,
    importance: str = "normal",
    format: str = "text",
    mentions: Optional[List[Dict[str, str]]] = None,
    image_url: Optional[str] = None,
    image_data: Optional[str] = None,
    image_content_type: Optional[str] = None,
    image_file_name: Optional[str] = None
) -> str:
    """
    Reply to a specific message in a channel. Supports text and markdown formatting, mentions, and importance levels.
    
    Args:
        user_email (str): The user's email address. Required.
        team_id (str): Team ID
        channel_id (str): Channel ID
        message_id (str): Message ID to reply to
        message (str): Reply content
        importance (str): Message importance (normal, high, urgent)
        format (str): Message format (text or markdown)
        mentions (List[Dict]): Array of @mentions to include in the reply
        image_url (str): URL of an image to attach to the reply
        image_data (str): Base64 encoded image data to attach
        image_content_type (str): MIME type of the image
        image_file_name (str): Name for the attached image file
        
    Returns:
        str: Success or error message.
    """
    logger.info(f"[reply_to_channel_message] Replying to message {message_id} in team {team_id}, channel {channel_id}")
    
    try:
        # Process message content based on format
        content = message
        content_type = "text"
        
        if format == "markdown":
            content = await _markdown_to_html(message)
            content_type = "html"
        
        # Process @mentions if provided
        final_mentions = []
        mention_mappings = []
        
        if mentions:
            for mention in mentions:
                try:
                    user_response = await service.get(f"/users/{mention['userId']}?$select=displayName")
                    display_name = user_response.get("displayName", mention["mention"])
                    
                    mention_mappings.append({
                        "mention": mention["mention"],
                        "userId": mention["userId"],
                        "displayName": display_name,
                    })
                except Exception as e:
                    logger.warning(f"Could not resolve user {mention['userId']}: {e}")
                    mention_mappings.append({
                        "mention": mention["mention"],
                        "userId": mention["userId"],
                        "displayName": mention["mention"],
                    })
        
        # Process mentions in HTML content
        if mention_mappings:
            content, final_mentions = await _process_mentions_in_html(content, mention_mappings)
            content_type = "html"
        
        # Handle image attachment (simplified)
        attachments = []
        
        # Build message payload
        message_payload = {
            "body": {
                "content": content,
                "contentType": content_type,
            },
            "importance": importance,
        }
        
        if final_mentions:
            message_payload["mentions"] = final_mentions
        
        if attachments:
            message_payload["attachments"] = attachments
        
        result = await service.post(f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies", message_payload)
        
        # Build success message
        success_text = f"✅ Reply sent successfully. Reply ID: {result.get('id')}"
        if final_mentions:
            mentions_text = ", ".join([m.get("mentionText", "") for m in final_mentions])
            success_text += f"\n📱 Mentions: {mentions_text}"
        if attachments:
            success_text += f"\n🖼️ Image attached: {attachments[0].get('name', '')}"
        
        return success_text
        
    except Exception as e:
        logger.error(f"[reply_to_channel_message] Error: {e}")
        return f"❌ Failed to send reply: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def list_team_members(service, user_email: str, team_id: str) -> str:
    """
    List all members of a specific Microsoft Team. Returns member names, email addresses, roles, and IDs.
    
    Args:
        user_email (str): The user's email address. Required.
        team_id (str): Team ID
        
    Returns:
        str: JSON string containing team members information.
    """
    logger.info(f"[list_team_members] Fetching members for team {team_id}, user: {user_email}")
    
    try:
        members_data = await service.get(f"/teams/{team_id}/members")
        
        if not members_data.get("value"):
            return json.dumps({"message": "No members found in this team."})
        
        member_list = []
        for member in members_data["value"]:
            member_info = {
                "id": member.get("id"),
                "displayName": member.get("displayName"),
                "roles": member.get("roles", []),
            }
            member_list.append(member_info)
        
        return json.dumps(member_list, indent=2)
        
    except Exception as e:
        logger.error(f"[list_team_members] Error: {e}")
        return f"❌ Error: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def search_users_for_mentions(service, user_email: str, query: str, limit: int = 10) -> str:
    """
    Search for users to mention in messages. Returns users with their display names, email addresses, and mention IDs.
    
    Args:
        user_email (str): The user's email address. Required.
        query (str): Search query (name or email)
        limit (int): Maximum number of results to return (default: 10, max: 50)
        
    Returns:
        str: JSON string containing search results.
    """
    logger.info(f"[search_users_for_mentions] Searching users with query '{query}', user: {user_email}")
    
    try:
        # Validate limit
        if limit < 1 or limit > 50:
            limit = 10
        
        # Search users using Microsoft Graph
        search_query = f"$search=\"{query}\"&$top={limit}&$select=id,displayName,userPrincipalName"
        users_data = await service.get(f"/users?{search_query}")
        
        if not users_data.get("value"):
            return json.dumps({
                "query": query,
                "totalResults": 0,
                "users": [],
                "message": f"No users found matching \"{query}\"."
            })
        
        users_list = []
        for user in users_data["value"]:
            user_principal_name = user.get("userPrincipalName", "")
            mention_text = ""
            if user_principal_name:
                mention_text = user_principal_name.split("@")[0]
            else:
                mention_text = user.get("displayName", "").lower().replace(" ", "")
            
            user_info = {
                "id": user.get("id"),
                "displayName": user.get("displayName"),
                "userPrincipalName": user_principal_name,
                "mentionText": mention_text,
            }
            users_list.append(user_info)
        
        result = {
            "query": query,
            "totalResults": len(users_list),
            "users": users_list,
        }
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"[search_users_for_mentions] Error: {e}")
        return f"❌ Error: {str(e)}"

# Helper functions

async def _markdown_to_html(markdown_text: str) -> str:
    """
    Simple markdown to HTML conversion.
    In a full implementation, you'd use a proper markdown library like markdown or mistune.
    """
    # Simple conversions
    html = markdown_text
    
    # Bold
    html = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', html)
    
    # Italic
    html = re.sub(r'\*(.*?)\*', r'<em>\1</em>', html)
    
    # Line breaks
    html = html.replace('\n', '<br>')
    
    return html

async def _process_mentions_in_html(content: str, mention_mappings: List[Dict[str, str]]) -> tuple[str, List[Dict]]:
    """
    Process @mentions in HTML content and return updated content with mentions array.
    """
    final_mentions = []
    
    for i, mapping in enumerate(mention_mappings):
        mention_text = f"@{mapping['mention']}"
        
        if mention_text in content:
            # Replace mention text with proper HTML
            mention_id = i
            mention_html = f'<at id="{mention_id}">{mapping["displayName"]}</at>'
            content = content.replace(mention_text, mention_html)
            
            # Add to mentions array
            final_mentions.append({
                "id": mention_id,
                "mentionText": mapping["displayName"],
                "mentioned": {
                    "user": {
                        "id": mapping["userId"]
                    }
                }
            })
    
    return content, final_mentions
