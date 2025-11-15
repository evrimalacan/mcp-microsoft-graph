import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { ConversationMember } from '../../client/graph.types.js';
import { graphService } from '../../services/graph.js';
import type { OptimizedChat } from '../tools.types.js';

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
      const chats = await client.searchChats({ searchTerm, memberName });

      const chatList: OptimizedChat[] = chats.map((chat) => ({
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
