import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { OptimizedChatMessage } from '../tools.types.js';

const schema = z.object({
  chatId: z.string().describe('Chat ID'),
  messageId: z.string().describe('Message ID'),
});

export const getMessageTool = (server: McpServer) => {
  server.registerTool(
    'get_message',
    {
      title: 'Get Message',
      description:
        'Get a specific message by ID. Returns full message content, sender, attachments, and reactions.',
      inputSchema: schema.shape,
    },
    async ({ chatId, messageId }) => {
      const client = await graphService.getSDKClient();
      const message = await client.api(`/chats/${chatId}/messages/${messageId}`).get();

      const result: OptimizedChatMessage = {
        id: message.id,
        content: message.body?.content || undefined,
        from: message.from?.user?.displayName || undefined,
        createdDateTime: message.createdDateTime || undefined,
        reactions: message.reactions || [],
        attachments: message.attachments?.length ? message.attachments : undefined,
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    },
  );
};
