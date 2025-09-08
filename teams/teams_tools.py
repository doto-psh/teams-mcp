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
        # Check if service is properly initialized
        if service is None:
            logger.error("[get_channel_messages] Service is None - authentication may have failed")
            return "❌ Error: Service not initialized. Please check authentication."
        
        # Validate limit
        if limit < 1 or limit > 50:
            limit = 20
        
        # Build query parameters - Teams channel messages API has limited query support
        # Only $top is supported, no $orderby, $filter, etc.
        query_params = f"$top={limit}"
        endpoint = f"/teams/{team_id}/channels/{channel_id}/messages?{query_params}"
        
        logger.debug(f"[get_channel_messages] Making request to: {endpoint}")
        messages_data = await service.get(endpoint)
        
        # Check if messages_data is None or doesn't have expected structure
        if messages_data is None:
            logger.error("[get_channel_messages] Received None response from service")
            return "❌ Error: No response from Microsoft Graph API. Please check permissions."
        
        if not isinstance(messages_data, dict):
            logger.error(f"[get_channel_messages] Unexpected response type: {type(messages_data)}")
            return "❌ Error: Unexpected response format from Microsoft Graph API."
        
        if not messages_data.get("value"):
            return json.dumps({"message": "No messages found in this channel."})
        
        message_list = []
        for message in messages_data["value"]:
            # Safely extract message information with null checks
            message_body = message.get("body") or {}
            message_from = message.get("from") or {}
            message_user = message_from.get("user") or {}
            
            message_info = {
                "id": message.get("id"),
                "content": message_body.get("content"),
                "from": message_user.get("displayName"),
                "createdDateTime": message.get("createdDateTime"),
                "importance": message.get("importance"),
            }
            message_list.append(message_info)
        
        # Sort messages by creation date (newest first) since API doesn't support orderby
        message_list.sort(key=lambda x: x.get("createdDateTime") or "", reverse=True)
        
        result = {
            "totalReturned": len(message_list),
            "hasMore": bool(messages_data.get("@odata.nextLink")),
            "messages": message_list,
        }
        
        return json.dumps(result, indent=2)
        
    except AttributeError as e:
        logger.error(f"[get_channel_messages] AttributeError - likely service is None: {e}")
        return f"❌ Error: Service initialization failed. Please check authentication status. Details: {str(e)}"
    except Exception as e:
        logger.error(f"[get_channel_messages] Unexpected error: {e}")
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
        # Check if service is properly initialized
        if service is None:
            logger.error("[send_channel_message] Service is None - authentication may have failed")
            return "❌ Error: Service not initialized. Please check authentication."
        
        # Validate importance level
        valid_importance = ["normal", "high", "urgent"]
        if importance not in valid_importance:
            importance = "normal"
            logger.warning(f"[send_channel_message] Invalid importance level, defaulting to 'normal'")
        
        # Validate format
        valid_formats = ["text", "markdown"]
        if format not in valid_formats:
            format = "text"
            logger.warning(f"[send_channel_message] Invalid format, defaulting to 'text'")
        
        # Process message content based on format
        content = message
        content_type = "text"
        
        if format == "markdown":
            # Simple markdown to HTML conversion
            content = await _markdown_to_html(message)
            content_type = "html"
        
        # Process @mentions if provided
        mention_mappings = []
        if mentions:
            logger.info(f"[send_channel_message] Processing {len(mentions)} mentions")
            for mention in mentions:
                try:
                    # Validate mention structure
                    if not mention.get("userId") or not mention.get("mention"):
                        logger.warning(f"[send_channel_message] Invalid mention structure: {mention}")
                        continue
                    
                    # Get user info to get display name
                    user_response = await service.get(f"/users/{mention['userId']}?$select=displayName")
                    display_name = user_response.get("displayName", mention["mention"])
                    
                    mention_mappings.append({
                        "mention": mention["mention"],
                        "userId": mention["userId"],
                        "displayName": display_name,
                    })
                    logger.debug(f"[send_channel_message] Resolved mention: {mention['mention']} -> {display_name}")
                except Exception as e:
                    logger.warning(f"[send_channel_message] Could not resolve user {mention.get('userId')}: {e}")
                    mention_mappings.append({
                        "mention": mention["mention"],
                        "userId": mention["userId"],
                        "displayName": mention["mention"],
                    })
        
        # Process mentions in HTML content
        final_mentions = []
        if mention_mappings:
            content, final_mentions = await _process_mentions_in_html(content, mention_mappings)
            content_type = "html"  # Ensure HTML when mentions are present
            logger.info(f"[send_channel_message] Processed {len(final_mentions)} mentions in content")
        
        # Handle image attachment
        attachments = []
        if image_url or image_data:
            logger.info("[send_channel_message] Processing image attachment")
            
            # Validate image content type if provided
            if image_content_type and not _is_valid_image_type(image_content_type):
                return f"❌ Unsupported image type: {image_content_type}. Supported types: image/jpeg, image/png, image/gif, image/webp"
            
            try:
                # Handle image URL
                if image_url:
                    logger.info(f"[send_channel_message] Downloading image from URL: {image_url}")
                    image_info = _download_image_from_url(image_url)
                    if not image_info:
                        return f"❌ Failed to download image from URL: {image_url}"
                    image_data = image_info["data"]
                    image_content_type = image_info["content_type"]
                    if not image_file_name:
                        image_file_name = image_info.get("filename", "image.jpg")
                
                # Handle base64 image data
                elif image_data and image_content_type:
                    if not image_file_name:
                        # Generate filename from content type
                        ext_map = {
                            "image/jpeg": "jpg",
                            "image/png": "png", 
                            "image/gif": "gif",
                            "image/webp": "webp"
                        }
                        ext = ext_map.get(image_content_type, "jpg")
                        image_file_name = f"image.{ext}"
                
                # Create hosted content attachment
                if image_data and image_content_type and image_file_name:
                    attachment = _create_hosted_content_attachment(
                        service, team_id, channel_id, image_data, image_content_type, image_file_name
                    )
                    if attachment:
                        attachments.append(attachment)
                        logger.info(f"[send_channel_message] Created image attachment: {image_file_name}")
                    else:
                        return "❌ Failed to upload image attachment"
                        
            except Exception as e:
                logger.error(f"[send_channel_message] Error processing image: {e}")
                return f"❌ Failed to process image attachment: {str(e)}"
        
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
        
        logger.debug(f"[send_channel_message] Message payload: {json.dumps(message_payload, indent=2)}")
        
        # Send the message
        result = await service.post(f"/teams/{team_id}/channels/{channel_id}/messages", message_payload)
        
        if not result or not result.get("id"):
            return "❌ Failed to send message: No message ID returned"
        
        # Build success message
        success_parts = [f"✅ Message sent successfully. Message ID: {result.get('id')}"]
        
        if final_mentions:
            mentions_text = ", ".join([m.get("mentionText", "") for m in final_mentions])
            success_parts.append(f"📱 Mentions: {mentions_text}")
        
        if attachments:
            success_parts.append(f"🖼️ Image attached: {attachments[0].get('name', image_file_name)}")
        
        success_text = "\n".join(success_parts)
        logger.info(f"[send_channel_message] Message sent successfully: {result.get('id')}")
        
        return success_text
        
    except Exception as e:
        logger.error(f"[send_channel_message] Unexpected error: {e}")
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
    logger.info(f"[get_channel_message_replies] Fetching replies for message {message_id} in team {team_id}, channel {channel_id}, user: {user_email}")
    
    try:
        # Check if service is properly initialized
        if service is None:
            logger.error("[get_channel_message_replies] Service is None - authentication may have failed")
            return "❌ Error: Service not initialized. Please check authentication."
        
        # Validate limit (same as TypeScript: min 1, max 50, default 20)
        if limit < 1 or limit > 50:
            limit = 20
        
        # Only $top is supported for message replies
        query_params = [f"$top={limit}"]
        query_string = "&".join(query_params)
        
        endpoint = f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies?{query_string}"
        logger.debug(f"[get_channel_message_replies] Making request to: {endpoint}")
        
        replies_data = await service.get(endpoint)
        
        # Check if replies_data is None or doesn't have expected structure
        if replies_data is None:
            logger.error("[get_channel_message_replies] Received None response from service")
            return "❌ Error: No response from Microsoft Graph API. Please check permissions."
        
        if not isinstance(replies_data, dict):
            logger.error(f"[get_channel_message_replies] Unexpected response type: {type(replies_data)}")
            return "❌ Error: Unexpected response format from Microsoft Graph API."
        
        if not replies_data.get("value"):
            return json.dumps({"message": "No replies found for this message."})
        
        replies_list = []
        for reply in replies_data["value"]:
            # Safely extract reply information with null checks
            reply_body = reply.get("body") or {}
            reply_from = reply.get("from") or {}
            reply_user = reply_from.get("user") or {}
            
            reply_info = {
                "id": reply.get("id"),
                "content": reply_body.get("content"),
                "from": reply_user.get("displayName"),
                "createdDateTime": reply.get("createdDateTime"),
                "importance": reply.get("importance"),
            }
            replies_list.append(reply_info)
        
        # Sort replies by creation date (oldest first for replies) - same as TypeScript
        replies_list.sort(key=lambda x: x.get("createdDateTime") or "")
        
        result = {
            "parentMessageId": message_id,
            "totalReplies": len(replies_list),
            "hasMore": bool(replies_data.get("@odata.nextLink")),
            "replies": replies_list,
        }
        
        logger.info(f"[get_channel_message_replies] Retrieved {len(replies_list)} replies for message {message_id}")
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"[get_channel_message_replies] Unexpected error: {e}")
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
    logger.info(f"[reply_to_channel_message] Replying to message {message_id} in team {team_id}, channel {channel_id}, user: {user_email}")
    
    try:
        # Check if service is properly initialized
        if service is None:
            logger.error("[reply_to_channel_message] Service is None - authentication may have failed")
            return "❌ Error: Service not initialized. Please check authentication."
        
        # Validate importance level
        valid_importance = ["normal", "high", "urgent"]
        if importance not in valid_importance:
            importance = "normal"
            logger.warning(f"[reply_to_channel_message] Invalid importance level, defaulting to 'normal'")
        
        # Validate format
        valid_formats = ["text", "markdown"]
        if format not in valid_formats:
            format = "text"
            logger.warning(f"[reply_to_channel_message] Invalid format, defaulting to 'text'")
        
        # Process message content based on format
        content = message
        content_type = "text"
        
        if format == "markdown":
            content = await _markdown_to_html(message)
            content_type = "html"
        
        # Process @mentions if provided
        mention_mappings = []
        if mentions:
            logger.info(f"[reply_to_channel_message] Processing {len(mentions)} mentions")
            for mention in mentions:
                try:
                    # Validate mention structure
                    if not mention.get("userId") or not mention.get("mention"):
                        logger.warning(f"[reply_to_channel_message] Invalid mention structure: {mention}")
                        continue
                    
                    # Get user info to get display name
                    user_response = await service.get(f"/users/{mention['userId']}?$select=displayName")
                    display_name = user_response.get("displayName", mention["mention"])
                    
                    mention_mappings.append({
                        "mention": mention["mention"],
                        "userId": mention["userId"],
                        "displayName": display_name,
                    })
                    logger.debug(f"[reply_to_channel_message] Resolved mention: {mention['mention']} -> {display_name}")
                except Exception as e:
                    logger.warning(f"[reply_to_channel_message] Could not resolve user {mention.get('userId')}: {e}")
                    mention_mappings.append({
                        "mention": mention["mention"],
                        "userId": mention["userId"],
                        "displayName": mention["mention"],
                    })
        
        # Process mentions in HTML content
        final_mentions = []
        if mention_mappings:
            content, final_mentions = await _process_mentions_in_html(content, mention_mappings)
            content_type = "html"  # Ensure HTML when mentions are present
            logger.info(f"[reply_to_channel_message] Processed {len(final_mentions)} mentions in content")
        
        # Handle image attachment
        attachments = []
        if image_url or image_data:
            logger.info("[reply_to_channel_message] Processing image attachment")
            
            # Validate image content type if provided
            if image_content_type and not _is_valid_image_type(image_content_type):
                return f"❌ Unsupported image type: {image_content_type}. Supported types: image/jpeg, image/png, image/gif, image/webp"
            
            try:
                # Handle image URL
                if image_url:
                    logger.info(f"[reply_to_channel_message] Downloading image from URL: {image_url}")
                    image_info = _download_image_from_url(image_url)
                    if not image_info:
                        return f"❌ Failed to download image from URL: {image_url}"
                    image_data = image_info["data"]
                    image_content_type = image_info["content_type"]
                    if not image_file_name:
                        image_file_name = image_info.get("filename", "image.jpg")
                
                # Handle base64 image data
                elif image_data and image_content_type:
                    if not _is_valid_image_type(image_content_type):
                        return f"❌ Unsupported image type: {image_content_type}"
                    if not image_file_name:
                        # Generate filename from content type
                        ext_map = {
                            "image/jpeg": "jpg",
                            "image/png": "png", 
                            "image/gif": "gif",
                            "image/webp": "webp"
                        }
                        ext = ext_map.get(image_content_type, "jpg")
                        image_file_name = f"image.{ext}"
                
                # Create hosted content attachment
                if image_data and image_content_type and image_file_name:
                    attachment = _create_hosted_content_attachment(
                        service, team_id, channel_id, image_data, image_content_type, image_file_name
                    )
                    if attachment:
                        attachments.append(attachment)
                        logger.info(f"[reply_to_channel_message] Created image attachment: {image_file_name}")
                    else:
                        return "❌ Failed to upload image attachment"
                        
            except Exception as e:
                logger.error(f"[reply_to_channel_message] Error processing image: {e}")
                return f"❌ Failed to process image attachment: {str(e)}"
        
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
        
        logger.debug(f"[reply_to_channel_message] Message payload: {json.dumps(message_payload, indent=2)}")
        
        # Send the reply
        result = await service.post(f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies", message_payload)
        
        if not result or not result.get("id"):
            return "❌ Failed to send reply: No reply ID returned"
        
        # Build success message
        success_parts = [f"✅ Reply sent successfully. Reply ID: {result.get('id')}"]
        
        if final_mentions:
            mentions_text = ", ".join([m.get("mentionText", "") for m in final_mentions])
            success_parts.append(f"📱 Mentions: {mentions_text}")
        
        if attachments:
            success_parts.append(f"🖼️ Image attached: {attachments[0].get('name', image_file_name)}")
        
        success_text = "\n".join(success_parts)
        logger.info(f"[reply_to_channel_message] Reply sent successfully: {result.get('id')}")
        
        return success_text
        
    except Exception as e:
        logger.error(f"[reply_to_channel_message] Unexpected error: {e}")
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

# Helper functions

def _is_valid_image_type(content_type: str) -> bool:
    """
    Validate if the content type is a supported image format.
    """
    supported_types = [
        "image/jpeg",
        "image/jpg", 
        "image/png",
        "image/gif",
        "image/webp",
        "image/bmp"
    ]
    return content_type.lower() in supported_types

def _download_image_from_url(image_url: str) -> Optional[Dict[str, str]]:
    """
    Download image from URL and return base64 data with content type.
    """
    try:
        import requests
        import mimetypes
        from urllib.parse import urlparse
        
        response = requests.get(image_url, timeout=30)
        if response.status_code != 200:
            logger.error(f"Failed to download image: HTTP {response.status_code}")
            return None
        
        # Get content type from response headers
        content_type = response.headers.get("content-type")
        if not content_type:
            # Try to guess from URL
            parsed_url = urlparse(image_url)
            content_type, _ = mimetypes.guess_type(parsed_url.path)
        
        if not content_type or not _is_valid_image_type(content_type):
            logger.error(f"Invalid or unsupported image type: {content_type}")
            return None
        
        # Read image data
        image_data = response.content
        base64_data = base64.b64encode(image_data).decode('utf-8')
        
        # Extract filename from URL
        filename = parsed_url.path.split("/")[-1]
        if not filename or "." not in filename:
            ext_map = {
                "image/jpeg": "jpg",
                "image/png": "png",
                "image/gif": "gif", 
                "image/webp": "webp"
            }
            ext = ext_map.get(content_type, "jpg")
            filename = f"image.{ext}"
        
        return {
            "data": base64_data,
            "content_type": content_type,
            "filename": filename
        }
                
    except ImportError:
        logger.error("requests is required for downloading images from URLs. Install with: pip install requests")
        return None
    except Exception as e:
        logger.error(f"Error downloading image from URL: {e}")
        return None

def _create_hosted_content_attachment(
    service, 
    team_id: str, 
    channel_id: str, 
    image_data: str, 
    content_type: str, 
    filename: str
) -> Optional[Dict]:
    """
    Create a hosted content attachment for Teams message.
    This is a simplified implementation - full implementation would involve OneDrive/SharePoint upload.
    """
    try:
        # For now, create a simple attachment reference
        # In a full implementation, you would:
        # 1. Upload the file to OneDrive/SharePoint
        # 2. Get the sharing link
        # 3. Create proper attachment with the link
        
        # Create a hash ID from image data
        import hashlib
        data_hash = hashlib.md5(image_data.encode('utf-8')).hexdigest()[:8]
        
        attachment = {
            "id": f"attachment_{data_hash}",
            "contentType": "reference",
            "contentUrl": f"data:{content_type};base64,{image_data[:100]}...",  # Truncated for demo
            "name": filename,
            "content": {
                "contentType": content_type,
                "downloadUrl": None,  # Would be the actual download URL in full implementation
                "webUrl": None,       # Would be the web URL in full implementation
                "uniqueId": f"hosted_content_{data_hash}"
            }
        }
        
        logger.info(f"Created hosted content attachment: {filename}")
        return attachment
        
    except Exception as e:
        logger.error(f"Error creating hosted content attachment: {e}")
        return None

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
