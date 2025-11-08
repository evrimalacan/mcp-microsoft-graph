import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { Chat, ChatSummary, ConversationMember, CreateChatPayload, User } from '../../types/graph.js';

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

      const me = (await client.api('/me').get()) as User;

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
        const user = (await client.api(`/users/${email}`).get()) as User;
        members.push({
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          user: {
            id: user?.id,
          },
          roles: ['guest'],
        } as ConversationMember);
      }

      const chatData: CreateChatPayload = {
        chatType: userEmails.length === 1 ? 'oneOnOne' : 'group',
        members,
      };

      if (topic && userEmails.length > 1) {
        chatData.topic = topic;
      }

      const newChat = (await client.api('/chats').post(chatData)) as Chat;

      const chatSummary: ChatSummary = {
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
