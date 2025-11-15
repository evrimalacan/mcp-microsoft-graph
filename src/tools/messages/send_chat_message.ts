import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  chatId: z.string().describe('Chat ID'),
  message: z.string().describe('Message content'),
  importance: z.enum(['normal', 'high', 'urgent']).optional().describe('Message importance'),
  format: z.enum(['text', 'markdown']).optional().describe('Message format (text or markdown)'),
  mentions: z
    .array(
      z.object({
        mention: z.string().describe("The @mention text (e.g., 'john.doe' or 'john.doe@company.com')"),
        userId: z.string().describe('Azure AD User ID of the mentioned user'),
      }),
    )
    .optional()
    .describe('Array of @mentions to include in the message'),
});

export const sendChatMessageTool = (server: McpServer) => {
  server.registerTool(
    'send_chat_message',
    {
      title: 'Send Chat Message',
      description:
        'Send a message to a specific chat conversation. Supports text and markdown formatting, mentions, and importance levels.',
      inputSchema: schema.shape,
    },
    async ({ chatId, message, importance, format, mentions }) => {
      const client = await graphService.getClient();

      const result = await client.sendChatMessage({
        chatId,
        message,
        importance,
        format,
        mentions,
      });

      return {
        content: [
          {
            type: 'text',
            text: `✅ Message sent successfully. Message ID: ${result.id}`,
          },
        ],
      };
    },
  );
};
