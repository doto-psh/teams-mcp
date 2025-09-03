import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { DeviceCodeCredential } from "@azure/identity";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphService } from "../services/graph.js";

const CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const TOKEN_PATH = join(homedir(), ".msgraph-mcp-auth.json");

export function registerAuthTools(server: McpServer, graphService: GraphService) {
  // Authentication status tool
  server.tool(
    "auth_status",
    "Check the authentication status of the Microsoft Graph connection. Returns whether the user is authenticated and shows their basic profile information.",
    {},
    async () => {
      const status = await graphService.getAuthStatus();
      return {
        content: [
          {
            type: "text",
            text: status.isAuthenticated
              ? `✅ Authenticated as ${status.displayName || "Unknown User"} (${status.userPrincipalName || "No email available"})`
              : "❌ Not authenticated. Use the 'authenticate' tool to authenticate.",
          },
        ],
      };
    }
  );

  // Microsoft Graph Authentication tool
  server.tool(
    "authenticate",
    "Authenticate with Microsoft Graph using device code flow. This will guide you through the authentication process.",
    {},
    async () => {
      try {
        // Use Promise to manage device code information and authentication completion
        return new Promise((resolve) => {
          let deviceCodeInfo = "";
          
          const credential = new DeviceCodeCredential({
            clientId: CLIENT_ID,
            tenantId: "common",
            userPromptCallback: (info) => {
              // Provide device code information to user immediately
              deviceCodeInfo = `🔐 Microsoft Device Code Authentication

📋 Code: ${info.userCode}
🌐 Authentication URL: ${info.verificationUri}

Please follow these steps:
1. Open ${info.verificationUri} in your browser
2. Enter the code '${info.userCode}'
3. Sign in with your Microsoft account
4. Wait for authentication to complete...

⏳ Authentication in progress...`;

              // Show device code information to user first
              resolve({
                content: [
                  {
                    type: "text",
                    text: deviceCodeInfo,
                  },
                ],
              });
            },
          });

          // Proceed with token acquisition in background
          credential.getToken([
            "User.Read",
            "User.ReadBasic.All",
            "Team.ReadBasic.All",
            "Channel.ReadBasic.All",
            "ChannelMessage.Read.All",
            "ChannelMessage.Send",
            "TeamMember.Read.All",
            "Chat.ReadBasic",
            "Chat.ReadWrite",
          ]).then(async (token) => {
            if (token) {
              // Save authentication info with the actual token
              const authInfo = {
                clientId: CLIENT_ID,
                authenticated: true,
                timestamp: new Date().toISOString(),
                expiresAt: token.expiresOnTimestamp
                  ? new Date(token.expiresOnTimestamp).toISOString()
                  : undefined,
                token: token.token,
              };

              await fs.writeFile(TOKEN_PATH, JSON.stringify(authInfo, null, 2));
              
              // No separate notification after authentication completion (device code info already provided)
            }
          }).catch((error) => {
            // Even if error occurs, device code information has already been provided
            console.error("Authentication error:", error);
          });
        });
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `❌ Authentication failed: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
        };
      }
    }
  );

  // Check authentication status tool
  server.tool(
    "check_auth",
    "Check the detailed authentication status including token expiration and user information.",
    {},
    async () => {
      try {
        const data = await fs.readFile(TOKEN_PATH, "utf8");
        const authInfo = JSON.parse(data);

        if (authInfo.authenticated && authInfo.clientId) {
          let message = `✅ Authentication found\n📅 Authenticated on: ${authInfo.timestamp}`;

          // Check if we have expiration info
          if (authInfo.expiresAt) {
            const expiresAt = new Date(authInfo.expiresAt);
            const now = new Date();

            if (expiresAt > now) {
              message += `\n⏰ Token expires: ${expiresAt.toLocaleString()}\n🎯 Ready to use with MCP server!`;
            } else {
              message += "\n⚠️ Token may have expired - please re-authenticate using the 'authenticate' tool";
            }
          } else {
            message += "\n🎯 Ready to use with MCP server!";
          }

          return {
            content: [
              {
                type: "text",
                text: message,
              },
            ],
          };
        } else {
          return {
            content: [
              {
                type: "text",
                text: "❌ Invalid authentication data found",
              },
            ],
          };
        }
      } catch (_error) {
        return {
          content: [
            {
              type: "text",
              text: "❌ No authentication found. Use the 'authenticate' tool to authenticate.",
            },
          ],
        };
      }
    }
  );

  // Logout tool
  server.tool(
    "logout",
    "Clear the stored authentication credentials. You will need to re-authenticate to use Microsoft Graph tools.",
    {},
    async () => {
      try {
        await fs.unlink(TOKEN_PATH);
        return {
          content: [
            {
              type: "text",
              text: "✅ Successfully logged out\n🔄 Use the 'authenticate' tool to re-authenticate",
            },
          ],
        };
      } catch (_error) {
        return {
          content: [
            {
              type: "text",
              text: "ℹ️ No authentication to clear",
            },
          ],
        };
      }
    }
  );
}
