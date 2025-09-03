"""
Microsoft Teams Chat Tools

This module provides MCP tools for interacting with Microsoft Teams Chat via Graph API.
"""

import json
import logging
from typing import List, Dict, Any, Optional
from datetime import datetime
from auth.service_decorator_teams import require_teams_service
from core.server import server

logger = logging.getLogger(__name__)

@server.tool()
@require_teams_service("teams", "teams_read")
async def list_chats(service, user_email: str) -> str:
    """
    List all recent chats (1:1 conversations and group chats) that the current user participates in. Returns chat topics, types, and participant information.
    
    Args:
        user_email (str): The user's email address. Required.
        
    Returns:
        str: JSON string containing chats information.
    """
    logger.info(f"[list_chats] Fetching chats for user: {user_email}")
    
    try:
        # Build query parameters
        query_params = "$expand=members"
        
        chats_data = await service.get(f"/me/chats?{query_params}")
        
        if not chats_data.get("value"):
            return json.dumps({"message": "No chats found."})
        
        chat_list = []
        for chat in chats_data["value"]:
            members_names = []
            if chat.get("members"):
                members_names = [member.get("displayName", "") for member in chat["members"] if member.get("displayName")]
            
            chat_info = {
                "id": chat.get("id"),
                "topic": chat.get("topic") or "No topic",
                "chatType": chat.get("chatType"),
                "members": ", ".join(members_names) if members_names else "No members",
            }
            chat_list.append(chat_info)
        
        return json.dumps(chat_list, indent=2)
        
    except Exception as e:
        logger.error(f"[list_chats] Error: {e}")
        return f"❌ Error: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_read")
async def get_chat_messages(
    service,
    user_email: str,
    chat_id: str,
    limit: int = 20,
    since: Optional[str] = None,
    until: Optional[str] = None,
    from_user: Optional[str] = None,
    order_by: str = "createdDateTime",
    descending: bool = True
) -> str:
    """
    Retrieve recent messages from a specific chat conversation. Returns message content, sender information, and timestamps.
    
    Args:
        user_email (str): The user's email address. Required.
        chat_id (str): Chat ID (e.g. 19:meeting_Njhi..j@thread.v2)
        limit (int): Number of messages to retrieve (default: 20, max: 50)
        since (str): Get messages since this ISO datetime
        until (str): Get messages until this ISO datetime
        from_user (str): Filter messages from specific user ID
        order_by (str): Sort order (createdDateTime or lastModifiedDateTime)
        descending (bool): Sort in descending order (newest first)
        
    Returns:
        str: JSON string containing messages information.
    """
    logger.info(f"[get_chat_messages] Fetching messages for chat {chat_id}, user: {user_email}")
    
    try:
        # Validate limit
        if limit < 1 or limit > 50:
            limit = 20
        
        # Validate order_by and descending combination
        if order_by in ["createdDateTime", "lastModifiedDateTime"] and not descending:
            return f"❌ Error: QueryOptions to order by '{order_by}' in 'Ascending' direction is not supported."
        
        # Build query parameters
        query_params = [f"$top={limit}"]
        
        # Add ordering - Graph API only supports descending order for datetime fields in chat messages
        sort_direction = "desc" if descending else "asc"
        query_params.append(f"$orderby={order_by} {sort_direction}")
        
        # Add filters (only user filter is supported reliably)
        filters = []
        if from_user:
            filters.append(f"from/user/id eq '{from_user}'")
        
        if filters:
            query_params.append(f"$filter={' and '.join(filters)}")
        
        query_string = "&".join(query_params)
        
        messages_data = await service.get(f"/me/chats/{chat_id}/messages?{query_string}")
        
        if not messages_data.get("value"):
            return json.dumps({"message": "No messages found in this chat with the specified filters."})
        
        # Apply client-side date filtering since server-side filtering is not supported
        filtered_messages = messages_data["value"]
        
        if since or until:
            new_filtered_messages = []
            for message in messages_data["value"]:
                if not message.get("createdDateTime"):
                    new_filtered_messages.append(message)
                    continue
                
                try:
                    message_date = datetime.fromisoformat(message["createdDateTime"].replace('Z', '+00:00'))
                    
                    if since:
                        since_date = datetime.fromisoformat(since.replace('Z', '+00:00'))
                        if message_date <= since_date:
                            continue
                    
                    if until:
                        until_date = datetime.fromisoformat(until.replace('Z', '+00:00'))
                        if message_date >= until_date:
                            continue
                    
                    new_filtered_messages.append(message)
                except (ValueError, AttributeError):
                    # If date parsing fails, include the message
                    new_filtered_messages.append(message)
            
            filtered_messages = new_filtered_messages
        
        message_list = []
        for message in filtered_messages:
            message_info = {
                "id": message.get("id"),
                "content": message.get("body", {}).get("content"),
                "from": message.get("from", {}).get("user", {}).get("displayName"),
                "createdDateTime": message.get("createdDateTime"),
            }
            message_list.append(message_info)
        
        result = {
            "filters": {
                "since": since,
                "until": until,
                "fromUser": from_user
            },
            "filteringMethod": "client-side" if (since or until) else "server-side",
            "totalReturned": len(message_list),
            "hasMore": bool(messages_data.get("@odata.nextLink")),
            "messages": message_list,
        }
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"[get_chat_messages] Error: {e}")
        return f"❌ Error: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_write")
async def send_chat_message(
    service,
    user_email: str,
    chat_id: str,
    message: str,
    importance: str = "normal",
    format: str = "text",
    mentions: Optional[List[Dict[str, str]]] = None
) -> str:
    """
    Send a message to a specific chat conversation. Supports text and markdown formatting, mentions, and importance levels.
    
    Args:
        user_email (str): The user's email address. Required.
        chat_id (str): Chat ID
        message (str): Message content
        importance (str): Message importance (normal, high, urgent)
        format (str): Message format (text or markdown)
        mentions (List[Dict]): Array of @mentions to include in the message
        
    Returns:
        str: Success or error message.
    """
    logger.info(f"[send_chat_message] Sending message to chat {chat_id}, user: {user_email}")
    
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
        
        result = await service.post(f"/me/chats/{chat_id}/messages", message_payload)
        
        # Build success message
        success_text = f"✅ Message sent successfully. Message ID: {result.get('id')}"
        if final_mentions:
            mentions_text = ", ".join([m.get("mentionText", "") for m in final_mentions])
            success_text += f"\n📱 Mentions: {mentions_text}"
        
        return success_text
        
    except Exception as e:
        logger.error(f"[send_chat_message] Error: {e}")
        return f"❌ Failed to send message: {str(e)}"

@server.tool()
@require_teams_service("teams", "teams_write")
async def create_chat(
    service,
    user_email: str,
    user_emails: List[str],
    topic: Optional[str] = None
) -> str:
    """
    Create a new chat conversation. Can be a 1:1 chat (with one other user) or a group chat (with multiple users). Group chats can optionally have a topic.
    
    Args:
        user_email (str): The user's email address. Required.
        user_emails (List[str]): Array of user email addresses to add to chat
        topic (str): Chat topic (for group chats)
        
    Returns:
        str: Success or error message with chat ID.
    """
    logger.info(f"[create_chat] Creating chat with users {user_emails}, user: {user_email}")
    
    try:
        # Get current user ID
        me = await service.get("/me")
        
        # Create members array
        members = [
            {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "user": {
                    "id": me.get("id")
                },
                "roles": ["owner"]
            }
        ]
        
        # Add other users as members
        for email in user_emails:
            try:
                user = await service.get(f"/users/{email}")
                members.append({
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "user": {
                        "id": user.get("id")
                    },
                    "roles": ["member"]
                })
            except Exception as e:
                logger.error(f"Could not find user {email}: {e}")
                return f"❌ Error: Could not find user {email}"
        
        chat_data = {
            "chatType": "oneOnOne" if len(user_emails) == 1 else "group",
            "members": members
        }
        
        if topic and len(user_emails) > 1:
            chat_data["topic"] = topic
        
        new_chat = await service.post("/chats", chat_data)
        
        return f"✅ Chat created successfully. Chat ID: {new_chat.get('id')}"
        
    except Exception as e:
        logger.error(f"[create_chat] Error: {e}")
        return f"❌ Error: {str(e)}"

# Helper functions (reused from teams_tools.py)

async def _markdown_to_html(markdown_text: str) -> str:
    """
    Simple markdown to HTML conversion.
    In a full implementation, you'd use a proper markdown library like markdown or mistune.
    """
    import re
    
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
