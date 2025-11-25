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
  chatTypes: z
    .array(z.enum(['oneOnOne', 'group', 'meeting']))
    .optional()
    .describe(
      'Filter by chat types. Options: "oneOnOne" (1:1 conversations), "group" (group chats), "meeting" (meeting chats). Can specify multiple types. If omitted, returns all chat types.',
    ),
  limit: z
    .number()
    .min(1)
    .max(50)
    .optional()
    .default(20)
    .describe('Maximum number of chats to return (1-50). Default is 20.'),
  skipToken: z
    .string()
    .optional()
    .describe(
      'Pagination token from previous response. Use the nextToken from a previous call to get the next page of results.',
    ),
});

export const searchChatsTool = (server: McpServer) => {
  server.registerTool(
    'search_chats',
    {
      title: 'Search Chats',
      description:
        'Search for chats (1:1 conversations, group chats, and meeting chats) that the current user participates in. Supports filtering by topic name, member name, and chat types. Returns chat topics, types, and participant information.',
      inputSchema: schema.shape,
    },
    async ({ searchTerm, memberName, chatTypes, limit, skipToken }) => {
      const client = await graphService.getClient();
      const result = await client.searchChats({ searchTerm, memberName, chatTypes, limit, skipToken });

      const chatList: OptimizedChat[] = result.chats.map((chat) => ({
        id: chat.id,
        topic: chat.topic,
        chatType: chat.chatType,
        members: chat.members?.map((member: ConversationMember & { email?: string }) => ({
          displayName: member.displayName,
          email: member.email,
        })),
      }));

      // Build response with pagination info
      const response: { chats: OptimizedChat[]; nextToken?: string } = {
        chats: chatList,
      };

      if (result.nextToken) {
        response.nextToken = result.nextToken;
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(response, null, 2),
          },
        ],
      };
    },
  );
};
