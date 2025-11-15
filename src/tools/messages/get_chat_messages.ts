import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { OptimizedChatMessage, OptimizedChatMessageWithFilters } from '../tools.types.js';

const schema = z.object({
  chatId: z.string().describe('Chat ID (e.g. 19:meeting_Njhi..j@thread.v2'),
  limit: z.number().min(1).max(50).optional().default(20).describe('Number of messages to retrieve'),
  from: z.string().optional().describe("Get messages from this ISO datetime (e.g., '2025-01-01T00:00:00Z')"),
  to: z.string().optional().describe("Get messages to this ISO datetime (e.g., '2025-01-31T23:59:59Z')"),
  fromUser: z.string().optional().describe('Filter messages from specific user ID'),
});

export const getChatMessagesTool = (server: McpServer) => {
  server.registerTool(
    'get_chat_messages',
    {
      title: 'Get Chat Messages',
      description:
        'Retrieve recent messages from a specific chat conversation. Returns message content, sender information, and timestamps.',
      inputSchema: schema.shape,
    },
    async ({ chatId, limit, from, to, fromUser }) => {
      const client = await graphService.getClient();
      const messages = await client.getChatMessages({ chatId, limit, from, to, fromUser });

      const messageList: OptimizedChatMessage[] = messages.map((message) => ({
        id: message.id,
        content: message.body?.content || undefined,
        from: message.from?.user?.displayName || undefined,
        createdDateTime: message.createdDateTime || undefined,
        reactions: message.reactions || [],
      }));

      const result: OptimizedChatMessageWithFilters = {
        filters: { from, to, fromUser },
        filteringMethod: from || to ? 'client-side' : 'server-side',
        totalReturned: messageList.length,
        hasMore: false, // We don't have access to @odata.nextLink after the client processes it
        messages: messageList,
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
