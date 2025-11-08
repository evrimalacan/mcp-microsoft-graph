import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { ChatMessage, GraphApiResponse, MessageSummary } from '../../types/graph.js';

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

      let query = client.api(`/me/chats/${chatId}/messages`).top(limit).orderby('createdDateTime desc');

      if (fromUser) {
        query = query.filter(`from/user/id eq '${fromUser}'`);
      }

      const response = (await query.get()) as GraphApiResponse<ChatMessage>;

      let filteredMessages = response?.value || [];

      if ((from || to) && filteredMessages.length > 0) {
        filteredMessages = filteredMessages.filter((message: ChatMessage) => {
          if (!message.createdDateTime) return true;

          const messageDate = new Date(message.createdDateTime);
          if (from) {
            const fromDate = new Date(from);
            if (messageDate <= fromDate) return false;
          }
          if (to) {
            const toDate = new Date(to);
            if (messageDate >= toDate) return false;
          }
          return true;
        });
      }

      const messageList = filteredMessages.map((message) => ({
        id: message.id,
        content: message.body?.content,
        from: message.from?.user?.displayName,
        createdDateTime: message.createdDateTime,
      }));

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                filters: { from, to, fromUser },
                filteringMethod: from || to ? 'client-side' : 'server-side',
                totalReturned: messageList.length,
                hasMore: !!response['@odata.nextLink'],
                messages: messageList,
              },
              null,
              2,
            ),
          },
        ],
      };
    },
  );
};
