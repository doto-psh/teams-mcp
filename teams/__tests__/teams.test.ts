import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  createMockMcpServer,
  mockChannel,
  mockChatMessage,
  mockConversationMember,
  mockTeam,
} from "../../test-utils/setup.js";
import type {
  Channel,
  ChatMessage,
  ConversationMember,
  GraphApiResponse,
  Team,
} from "../../types/graph.js";
import { registerTeamsTools } from "../teams.js";

describe("Teams Tools", () => {
  let mockServer: any;
  let mockGraphService: any;
  let mockClient: any;

  beforeEach(() => {
    mockServer = createMockMcpServer();
    mockClient = {
      api: vi.fn().mockReturnValue({
        get: vi.fn(),
        post: vi.fn(),
      }),
    };

    mockGraphService = {
      getClient: vi.fn().mockResolvedValue(mockClient),
    };

    vi.clearAllMocks();
  });

  describe("list_teams tool", () => {
    it("should register list_teams tool correctly", () => {
      registerTeamsTools(mockServer, mockGraphService);

      expect(mockServer.tool).toHaveBeenCalledWith(
        "list_teams",
        "List all Microsoft Teams that the current user is a member of. Returns team names, descriptions, and IDs.",
        {},
        expect.any(Function)
      );
    });

    it("should return list of joined teams", async () => {
      const teamsResponse: GraphApiResponse<Team> = {
        value: [mockTeam],
      };

      mockClient.api().get.mockResolvedValue(teamsResponse);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("list_teams");
      const result = await tool.handler();

      expect(mockClient.api).toHaveBeenCalledWith("/me/joinedTeams");
      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: JSON.stringify(
              [
                {
                  id: mockTeam.id,
                  displayName: mockTeam.displayName,
                  description: mockTeam.description,
                  isArchived: mockTeam.isArchived,
                },
              ],
              null,
              2
            ),
          },
        ],
      });
    });

    it("should handle empty teams list", async () => {
      const emptyResponse: GraphApiResponse<Team> = {
        value: [],
      };

      mockClient.api().get.mockResolvedValue(emptyResponse);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("list_teams");
      const result = await tool.handler();

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "No teams found.",
          },
        ],
      });
    });

    it("should handle API errors", async () => {
      mockClient.api().get.mockRejectedValue(new Error("Teams API error"));
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("list_teams");
      const result = await tool.handler();

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: "❌ Error: Teams API error",
          },
        ],
      });
    });
  });

  describe("list_channels tool", () => {
    it("should register list_channels tool with correct schema", () => {
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("list_channels");
      expect(tool).toBeDefined();
      expect(tool.schema.teamId).toBeDefined();
    });

    it("should list channels for a team", async () => {
      const channelsResponse: GraphApiResponse<Channel> = {
        value: [mockChannel],
      };

      mockClient.api().get.mockResolvedValue(channelsResponse);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("list_channels");
      const result = await tool.handler({ teamId: "test-team-id" });

      expect(mockClient.api).toHaveBeenCalledWith("/teams/test-team-id/channels");
      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: JSON.stringify(
              [
                {
                  id: mockChannel.id,
                  displayName: mockChannel.displayName,
                  description: mockChannel.description,
                  membershipType: mockChannel.membershipType,
                },
              ],
              null,
              2
            ),
          },
        ],
      });
    });

    it("should handle empty channels list", async () => {
      mockClient.api().get.mockResolvedValue({ value: [] });
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("list_channels");
      const result = await tool.handler({ teamId: "test-team-id" });

      expect(result.content[0].text).toBe("No channels found in this team.");
    });
  });

  describe("get_channel_messages tool", () => {
    it("should register get_channel_messages tool with correct schema", () => {
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_channel_messages");
      expect(tool).toBeDefined();
      expect(tool.schema.teamId).toBeDefined();
      expect(tool.schema.channelId).toBeDefined();
      expect(tool.schema.limit).toBeDefined();
    });

    it("should get channel messages with default limit", async () => {
      const messagesResponse: GraphApiResponse<ChatMessage> = {
        value: [mockChatMessage],
      };

      mockClient.api().get.mockResolvedValue(messagesResponse);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_channel_messages");
      const result = await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        limit: 20, // Explicitly pass the default limit
      });

      expect(mockClient.api).toHaveBeenCalledWith(
        "/teams/test-team-id/channels/test-channel-id/messages?$top=20"
      );

      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                totalReturned: 1,
                hasMore: false,
                messages: [
                  {
                    id: mockChatMessage.id,
                    content: mockChatMessage.body?.content,
                    from: mockChatMessage.from?.user?.displayName,
                    createdDateTime: mockChatMessage.createdDateTime,
                    importance: mockChatMessage.importance,
                  },
                ],
              },
              null,
              2
            ),
          },
        ],
      });
    });

    it("should get channel messages with custom limit", async () => {
      const messagesResponse: GraphApiResponse<ChatMessage> = {
        value: [mockChatMessage],
      };

      mockClient.api().get.mockResolvedValue(messagesResponse);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_channel_messages");
      await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        limit: 50,
      });

      expect(mockClient.api).toHaveBeenCalledWith(
        "/teams/test-team-id/channels/test-channel-id/messages?$top=50"
      );
    });

    it("should handle empty messages", async () => {
      mockClient.api().get.mockResolvedValue({ value: [] });
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_channel_messages");
      const result = await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
      });

      expect(result.content[0].text).toBe("No messages found in this channel.");
    });

    it("should sort messages by creation date (newest first)", async () => {
      const message1 = { ...mockChatMessage, id: "msg1", createdDateTime: "2024-01-01T10:00:00Z" };
      const message2 = { ...mockChatMessage, id: "msg2", createdDateTime: "2024-01-01T12:00:00Z" };
      const message3 = { ...mockChatMessage, id: "msg3", createdDateTime: "2024-01-01T11:00:00Z" };

      const messagesResponse: GraphApiResponse<ChatMessage> = {
        value: [message1, message2, message3], // Unsorted
      };

      mockClient.api().get.mockResolvedValue(messagesResponse);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_channel_messages");
      const result = await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
      });

      const response = JSON.parse(result.content[0].text);
      expect(response.messages[0].id).toBe("msg2"); // Newest first
      expect(response.messages[1].id).toBe("msg3");
      expect(response.messages[2].id).toBe("msg1"); // Oldest last
    });
  });

  describe("send_channel_message tool", () => {
    it("should register send_channel_message tool with correct schema", () => {
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("send_channel_message");
      expect(tool).toBeDefined();
      expect(tool.schema.teamId).toBeDefined();
      expect(tool.schema.channelId).toBeDefined();
      expect(tool.schema.message).toBeDefined();
      expect(tool.schema.importance).toBeDefined();
    });

    it("should send message with markdown format", async () => {
      const sentMessage = { ...mockChatMessage, id: "markdown-message-id" };
      mockClient.api().post.mockResolvedValue(sentMessage);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("send_channel_message");
      await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        message: "**Bold** _Italic_",
        format: "markdown",
      });

      expect(mockClient.api().post).toHaveBeenCalledWith({
        body: {
          content: expect.stringContaining("<strong>Bold</strong>"),
          contentType: "html",
        },
        importance: "normal",
      });
    });

    it("should send message with text format (default)", async () => {
      const sentMessage = { ...mockChatMessage, id: "text-message-id" };
      mockClient.api().post.mockResolvedValue(sentMessage);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("send_channel_message");
      await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        message: "Plain text message",
      });

      expect(mockClient.api().post).toHaveBeenCalledWith({
        body: {
          content: "Plain text message",
          contentType: "text",
        },
        importance: "normal",
      });
    });

    it("should send message with custom importance", async () => {
      const sentMessage = { ...mockChatMessage, id: "new-message-id" };
      mockClient.api().post.mockResolvedValue(sentMessage);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("send_channel_message");
      await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        message: "Urgent update!",
        importance: "urgent",
      });

      expect(mockClient.api().post).toHaveBeenCalledWith({
        body: {
          content: "Urgent update!",
          contentType: "text",
        },
        importance: "urgent",
      });
    });

    it("should handle send message errors", async () => {
      mockClient.api().post.mockRejectedValue(new Error("Send failed"));
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("send_channel_message");
      const result = await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        message: "Test message",
      });

      expect(result.content[0].text).toContain("❌ Failed to send message: Send failed");
    });

    it("should send reply with markdown format", async () => {
      const sentReply = { ...mockChatMessage, id: "markdown-reply-id" };
      mockClient.api().post.mockResolvedValue(sentReply);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("reply_to_channel_message");
      await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        messageId: "message-id",
        message: "**Bold** reply",
        format: "markdown",
      });

      expect(mockClient.api().post).toHaveBeenCalledWith({
        body: {
          content: expect.stringContaining("<strong>Bold</strong>"),
          contentType: "html",
        },
        importance: "normal",
      });
    });
  });

  describe("get_channel_message_replies tool", () => {
    it("should register get_channel_message_replies tool with correct schema", () => {
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_channel_message_replies");
      expect(tool).toBeDefined();
      expect(tool.schema.teamId).toBeDefined();
      expect(tool.schema.channelId).toBeDefined();
      expect(tool.schema.messageId).toBeDefined();
      expect(tool.schema.limit).toBeDefined();
    });

    it("should get message replies", async () => {
      const repliesResponse: GraphApiResponse<ChatMessage> = {
        value: [mockChatMessage],
      };

      mockClient.api().get.mockResolvedValue(repliesResponse);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_channel_message_replies");
      const result = await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        messageId: "test-message-id",
        limit: 10,
      });

      expect(mockClient.api).toHaveBeenCalledWith(
        "/teams/test-team-id/channels/test-channel-id/messages/test-message-id/replies?$top=10"
      );

      const response = JSON.parse(result.content[0].text);
      expect(response.parentMessageId).toBe("test-message-id");
      expect(response.totalReplies).toBe(1);
      expect(response.replies).toHaveLength(1);
    });

    it("should handle no replies found", async () => {
      mockClient.api().get.mockResolvedValue({ value: [] });
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_channel_message_replies");
      const result = await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        messageId: "test-message-id",
      });

      expect(result.content[0].text).toBe("No replies found for this message.");
    });

    it("should handle get replies errors", async () => {
      mockClient.api().get.mockRejectedValue(new Error("Message not found"));
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_channel_message_replies");
      const result = await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        messageId: "invalid-message-id",
      });

      expect(result.content[0].text).toContain("❌ Error: Message not found");
    });
  });

  describe("reply_to_channel_message tool", () => {
    it("should register reply_to_channel_message tool with correct schema", () => {
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("reply_to_channel_message");
      expect(tool).toBeDefined();
      expect(tool.schema.teamId).toBeDefined();
      expect(tool.schema.channelId).toBeDefined();
      expect(tool.schema.messageId).toBeDefined();
      expect(tool.schema.message).toBeDefined();
      expect(tool.schema.importance).toBeDefined();
    });

    it("should reply to a message with default importance", async () => {
      mockClient.api().post.mockResolvedValue({ id: "reply-123" });
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("reply_to_channel_message");
      const result = await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        messageId: "test-message-id",
        message: "This is a reply",
      });

      expect(mockClient.api).toHaveBeenCalledWith(
        "/teams/test-team-id/channels/test-channel-id/messages/test-message-id/replies"
      );
      expect(mockClient.api().post).toHaveBeenCalledWith({
        body: {
          content: "This is a reply",
          contentType: "text",
        },
        importance: "normal",
      });
      expect(result.content[0].text).toBe("✅ Reply sent successfully. Reply ID: reply-123");
    });

    it("should reply to a message with custom importance", async () => {
      mockClient.api().post.mockResolvedValue({ id: "reply-456" });
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("reply_to_channel_message");
      const _result = await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        messageId: "test-message-id",
        message: "Urgent reply!",
        importance: "urgent",
      });

      expect(mockClient.api().post).toHaveBeenCalledWith({
        body: {
          content: "Urgent reply!",
          contentType: "text",
        },
        importance: "urgent",
      });
    });

    it("should handle reply errors", async () => {
      mockClient.api().post.mockRejectedValue(new Error("Reply failed"));
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("reply_to_channel_message");
      const result = await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        messageId: "test-message-id",
        message: "Test reply",
      });

      expect(result.content[0].text).toContain("❌ Failed to send reply: Reply failed");
    });

    it("should reply with markdown format", async () => {
      mockClient.api().post.mockResolvedValue({ id: "reply-md" });
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("reply_to_channel_message");
      await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        messageId: "test-message-id",
        message: "**Reply** _Markdown_",
        format: "markdown",
      });

      expect(mockClient.api().post).toHaveBeenCalledWith({
        body: {
          content: expect.stringContaining("<strong>Reply</strong>"),
          contentType: "html",
        },
        importance: "normal",
      });
    });

    it("should reply with text format (default)", async () => {
      const sentReply = { ...mockChatMessage, id: "text-reply-id" };
      mockClient.api().post.mockResolvedValue(sentReply);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("reply_to_channel_message");
      await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        messageId: "message-id",
        message: "Plain text reply",
      });

      expect(mockClient.api().post).toHaveBeenCalledWith({
        body: {
          content: "Plain text reply",
          contentType: "text",
        },
        importance: "normal",
      });
    });

    it("should fallback to text for invalid format in reply", async () => {
      mockClient.api().post.mockResolvedValue({ id: "reply-fallback" });
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("reply_to_channel_message");
      await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        messageId: "test-message-id",
        message: "Fallback reply",
        format: "invalid-format",
      });

      expect(mockClient.api().post).toHaveBeenCalledWith({
        body: {
          content: "Fallback reply",
          contentType: "text",
        },
        importance: "normal",
      });
    });
  });

  describe("list_team_members tool", () => {
    it("should register list_team_members tool with correct schema", () => {
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("list_team_members");
      expect(tool).toBeDefined();
      expect(tool.schema.teamId).toBeDefined();
    });

    it("should list team members", async () => {
      const membersResponse: GraphApiResponse<ConversationMember> = {
        value: [mockConversationMember],
      };

      mockClient.api().get.mockResolvedValue(membersResponse);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("list_team_members");
      const result = await tool.handler({ teamId: "test-team-id" });

      expect(mockClient.api).toHaveBeenCalledWith("/teams/test-team-id/members");
      expect(result).toEqual({
        content: [
          {
            type: "text",
            text: JSON.stringify(
              [
                {
                  id: mockConversationMember.id,
                  displayName: mockConversationMember.displayName,
                  roles: mockConversationMember.roles,
                },
              ],
              null,
              2
            ),
          },
        ],
      });
    });

    it("should handle empty members list", async () => {
      mockClient.api().get.mockResolvedValue({ value: [] });
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("list_team_members");
      const result = await tool.handler({ teamId: "test-team-id" });

      expect(result.content[0].text).toBe("No members found in this team.");
    });

    it("should handle list members errors", async () => {
      mockClient.api().get.mockRejectedValue(new Error("Team not found"));
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("list_team_members");
      const result = await tool.handler({ teamId: "invalid-team-id" });

      expect(result.content[0].text).toContain("❌ Error: Team not found");
    });
  });

  describe("error handling", () => {
    it("should handle authentication errors in all tools", async () => {
      const authError = new Error("Not authenticated");
      mockGraphService.getClient.mockRejectedValue(authError);
      registerTeamsTools(mockServer, mockGraphService);

      const testCases = [
        { tool: "list_teams", params: {}, expectedError: "❌ Error: Not authenticated" },
        {
          tool: "list_channels",
          params: { teamId: "test" },
          expectedError: "❌ Error: Not authenticated",
        },
        {
          tool: "get_channel_messages",
          params: { teamId: "test", channelId: "test" },
          expectedError: "❌ Error: Not authenticated",
        },
        {
          tool: "send_channel_message",
          params: { teamId: "test", channelId: "test", message: "test" },
          expectedError: "❌ Failed to send message: Not authenticated",
        },
        {
          tool: "list_team_members",
          params: { teamId: "test" },
          expectedError: "❌ Error: Not authenticated",
        },
      ];

      for (const { tool: toolName, params, expectedError } of testCases) {
        const tool = mockServer.getTool(toolName);
        const result = await tool.handler(params);
        expect(result.content[0].text).toContain(expectedError);
      }
    });

    it("should handle unknown errors", async () => {
      mockClient.api().get.mockRejectedValue("String error");
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("list_teams");
      const result = await tool.handler();

      expect(result.content[0].text).toBe("❌ Error: Unknown error occurred");
    });
  });

  describe("input validation", () => {
    it("should handle invalid team IDs", async () => {
      mockClient.api().get.mockRejectedValue(new Error("Team not found"));
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("list_channels");
      const result = await tool.handler({ teamId: "invalid-team-id" });

      expect(result.content[0].text).toContain("❌ Error: Team not found");
    });

    it("should handle invalid channel IDs", async () => {
      mockClient.api().get.mockRejectedValue(new Error("Channel not found"));
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_channel_messages");
      const result = await tool.handler({
        teamId: "test-team-id",
        channelId: "invalid-channel-id",
      });

      expect(result.content[0].text).toContain("❌ Error: Channel not found");
    });

    it("should handle empty message content", async () => {
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("send_channel_message");
      const _result = await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
        message: "",
      });

      expect(mockClient.api().post).toHaveBeenCalledWith({
        body: {
          content: "",
          contentType: "text",
        },
        importance: "normal",
      });
    });
  });

  describe("message content handling", () => {
    it("should handle messages with missing body content", async () => {
      const messageWithoutBody = { ...mockChatMessage, body: undefined };
      const messagesResponse: GraphApiResponse<ChatMessage> = {
        value: [messageWithoutBody],
      };

      mockClient.api().get.mockResolvedValue(messagesResponse);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_channel_messages");
      const result = await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
      });

      const response = JSON.parse(result.content[0].text);
      expect(response.messages[0].content).toBeUndefined();
    });

    it("should handle messages with missing from user", async () => {
      const messageWithoutFrom = { ...mockChatMessage, from: undefined };
      const messagesResponse: GraphApiResponse<ChatMessage> = {
        value: [messageWithoutFrom],
      };

      mockClient.api().get.mockResolvedValue(messagesResponse);
      registerTeamsTools(mockServer, mockGraphService);

      const tool = mockServer.getTool("get_channel_messages");
      const result = await tool.handler({
        teamId: "test-team-id",
        channelId: "test-channel-id",
      });

      const response = JSON.parse(result.content[0].text);
      expect(response.messages[0].from).toBeUndefined();
    });
  });
});
