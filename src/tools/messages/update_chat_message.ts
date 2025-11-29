import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  chatId: z.string().describe('Chat ID'),
  messageId: z.string().describe('Message ID to update'),
  message: z.string().describe('New message content. Use @email to mention users (e.g., @john.doe@company.com).'),
  importance: z.enum(['normal', 'high', 'urgent']).optional().describe('Message importance level'),
  format: z.enum(['text', 'markdown']).optional().describe('Message format (text or markdown)'),
});

export const updateChatMessageTool = (server: McpServer) => {
  server.registerTool(
    'update_chat_message',
    {
      title: 'Update Chat Message',
      description: 'Update an existing chat message. Use @email to mention users.',
      inputSchema: schema.shape,
    },
    async ({ chatId, messageId, message, importance, format }) => {
      const client = await graphService.getClient();

      await client.updateChatMessage({
        chatId,
        messageId,
        message,
        importance,
        format,
      });

      return {
        content: [{ type: 'text', text: 'Message updated.' }],
      };
    },
  );
};
