# Microsoft Teams MCP Server - 도구 가이드

## 🧰 Available Tools
Note: All tools support automatic authentication via @require_teams_service() decorators with OAuth 2.1 flow and session caching.

## 🔐 Authentication (auth_tools.py)
| Tool | Description |
|------|-------------|
| start_teams_auth | Initiate Microsoft Teams OAuth authentication flow |
| logout_teams_auth | Clear stored credentials and logout from Teams session |

## 👥 Team Management (teams_tools.py)
| Tool | Description |
|------|-------------|
| list_teams | List all Teams that the user is a member of |
| list_channels | List all channels in a specific Team |
| get_channel_messages | Retrieve recent messages from a channel with filtering |
| send_channel_message | Send messages to channels with markdown, mentions, and importance levels |
| get_channel_message_replies | Get all replies to a specific channel message |
| reply_to_channel_message | Reply to specific messages in channels |
| list_team_members | List all members of a Team with roles and information |
| search_users_for_mentions | Search users for @mention functionality |

## 💬 Chat Management (chat_tools.py)
| Tool | Description |
|------|-------------|
| list_chats | List all chat conversations (1:1 and group chats) |
| get_chat_messages | Retrieve messages from chats with advanced filtering |
| send_chat_message | Send messages to chat conversations |
| create_chat | Create new chat conversations (1:1 or group) |

## 🔍 Search & Discovery (search_tools.py)
| Tool | Description |
|------|-------------|
| search_messages | Advanced message search across Teams using KQL syntax |
| get_recent_messages | Get recent messages with advanced filtering options |
| get_my_mentions | Find all messages where current user was @mentioned |

## 👤 User Management (users_tools.py)
| Tool | Description |
|------|-------------|
| get_current_user | Get current authenticated user's profile information |
| search_users | Search for users in the organization by name or email |
| get_user | Get detailed information about a specific user |
| get_user_manager | Get manager information for a user |
| get_user_direct_reports | Get direct reports of a user |
| get_user_photo | Get user profile photo information |
| get_organization_users | List users in the organization |
| search_users_advanced | Advanced user search with multiple criteria |
