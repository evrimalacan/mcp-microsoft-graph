import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { Chat, ChatSummary, ConversationMember, GraphApiResponse } from '../../types/graph.js';

const schema = z.object({
  searchTerm: z.string().optional().describe('Search for chats by topic name'),
  memberName: z
    .string()
    .optional()
    .describe(
      'Filter chats by member display name (case-sensitive). For best results, use the full display name (e.g., "John Smith" instead of "John")',
    ),
});

export const searchChatsTool = (server: McpServer) => {
  server.registerTool(
    'search_chats',
    {
      title: 'Search Chats',
      description:
        'Search for chats (1:1 conversations and group chats) that the current user participates in. Supports filtering by topic name and member name. Returns chat topics, types, and participant information.',
      inputSchema: schema.shape,
    },
    async ({ searchTerm, memberName }) => {
      const client = await graphService.getClient();

      let query = client.api(`/me/chats`).expand('members');

      const filters: string[] = [];

      if (searchTerm) {
        filters.push(`contains(topic, '${searchTerm}')`);
      }

      if (memberName) {
        filters.push(`members/any(c:contains(c/displayName, '${memberName}'))`);
      }

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const response = (await query.get()) as GraphApiResponse<Chat>;

      const chatList: ChatSummary[] = (response.value || []).map((chat) => ({
        id: chat.id,
        topic: chat.topic,
        chatType: chat.chatType,
        members: chat.members?.map((member: ConversationMember) => member.displayName).join(', '),
      }));

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(chatList, null, 2),
          },
        ],
      };
    },
  );
};
