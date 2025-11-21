import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { ConversationMember } from '../../client/graph.types.js';
import { graphService } from '../../services/graph.js';
import type { OptimizedChat } from '../tools.types.js';

const schema = z.object({
  chatId: z.string().describe('Chat ID (e.g. 19:meeting_Njhi..j@thread.v2)'),
});

export const getChatTool = (server: McpServer) => {
  server.registerTool(
    'get_chat',
    {
      title: 'Get Chat',
      description: 'Get details of a specific chat by its ID. Returns chat topic, type, and participant information.',
      inputSchema: schema.shape,
    },
    async ({ chatId }) => {
      const client = await graphService.getClient();
      const chat = await client.getChat({ chatId });

      const optimized: OptimizedChat = {
        id: chat.id,
        topic: chat.topic,
        chatType: chat.chatType,
        members: chat.members?.map((member: ConversationMember & { email?: string }) => ({
          displayName: member.displayName,
          email: member.email,
        })),
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(optimized, null, 2),
          },
        ],
      };
    },
  );
};
