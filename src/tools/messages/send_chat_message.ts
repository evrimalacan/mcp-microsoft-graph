import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { ChatMessage } from '../../types/graph.js';
import { markdownToHtml } from '../../utils/markdown.js';
import { processMentionsInHtml } from '../../utils/users.js';

const schema = z.object({
  chatId: z.string().describe('Chat ID'),
  message: z.string().describe('Message content'),
  importance: z.enum(['normal', 'high', 'urgent']).optional().describe('Message importance'),
  format: z.enum(['text', 'markdown']).optional().describe('Message format (text or markdown)'),
  mentions: z
    .array(
      z.object({
        mention: z.string().describe("The @mention text (e.g., 'john.doe' or 'john.doe@company.com')"),
        userId: z.string().describe('Azure AD User ID of the mentioned user'),
      }),
    )
    .optional()
    .describe('Array of @mentions to include in the message'),
});

export const sendChatMessageTool = (server: McpServer) => {
  server.registerTool(
    'send_chat_message',
    {
      title: 'Send Chat Message',
      description:
        'Send a message to a specific chat conversation. Supports text and markdown formatting, mentions, and importance levels.',
      inputSchema: schema.shape,
    },
    async ({ chatId, message, importance = 'normal', format = 'text', mentions }) => {
      const client = await graphService.getClient();

      let content: string;
      let contentType: 'text' | 'html';

      if (format === 'markdown') {
        content = await markdownToHtml(message);
        contentType = 'html';
      } else {
        content = message;
        contentType = 'text';
      }

      const mentionMappings: Array<{ mention: string; userId: string; displayName: string }> = [];
      if (mentions && mentions.length > 0) {
        for (const mention of mentions) {
          try {
            const userResponse = await client.api(`/users/${mention.userId}`).select('displayName').get();
            mentionMappings.push({
              mention: mention.mention,
              userId: mention.userId,
              displayName: userResponse.displayName || mention.mention,
            });
          } catch (_error) {
            mentionMappings.push({
              mention: mention.mention,
              userId: mention.userId,
              displayName: mention.mention,
            });
          }
        }
      }

      let finalMentions: Array<{
        id: number;
        mentionText: string;
        mentioned: { user: { id: string } };
      }> = [];
      if (mentionMappings.length > 0) {
        const result = processMentionsInHtml(content, mentionMappings);
        content = result.content;
        finalMentions = result.mentions;
        contentType = 'html';
      }

      const messagePayload: any = {
        body: {
          content,
          contentType,
        },
        importance,
      };

      if (finalMentions.length > 0) {
        messagePayload.mentions = finalMentions;
      }

      const result = (await client.api(`/me/chats/${chatId}/messages`).post(messagePayload)) as ChatMessage;

      const successText = `✅ Message sent successfully. Message ID: ${result.id}${
        finalMentions.length > 0 ? `\n📱 Mentions: ${finalMentions.map((m) => m.mentionText).join(', ')}` : ''
      }`;

      return {
        content: [
          {
            type: 'text',
            text: successText,
          },
        ],
      };
    },
  );
};
