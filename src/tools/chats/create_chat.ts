import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { ConversationMember } from '../../client/graph.types.js';
import { graphService } from '../../services/graph.js';
import type { OptimizedChat } from '../tools.types.js';

const schema = z.object({
  userEmails: z.array(z.string()).describe('Array of user email addresses to add to chat'),
  topic: z.string().optional().describe('Chat topic (for group chats)'),
});

export const createChatTool = (server: McpServer) => {
  server.registerTool(
    'create_chat',
    {
      title: 'Create Chat',
      description:
        'Create a new chat conversation. Can be a 1:1 chat (with one other user) or a group chat (with multiple users). Group chats can optionally have a topic.',
      inputSchema: schema.shape,
    },
    async ({ userEmails, topic }) => {
      const client = await graphService.getClient();

      const me = await client.getCurrentUser();

      const members: ConversationMember[] = [
        {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          user: {
            id: me?.id,
          },
          roles: ['owner'],
        } as ConversationMember,
      ];

      for (const email of userEmails) {
        const user = await client.getUser({ userId: email });
        members.push({
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          user: {
            id: user?.id,
          },
          roles: ['guest'],
        } as ConversationMember);
      }

      const chatType = userEmails.length === 1 ? 'oneOnOne' : 'group';
      const newChat = await client.createChat({
        chatType,
        members,
        ...(topic && userEmails.length > 1 && { topic }),
      });

      const chatSummary: OptimizedChat = {
        id: newChat.id,
        topic: newChat.topic,
        chatType: newChat.chatType,
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(chatSummary, null, 2),
          },
        ],
      };
    },
  );
};
