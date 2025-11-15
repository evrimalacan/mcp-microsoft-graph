import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  chatId: z.string().describe('Chat ID (e.g. 19:meeting_Njhi..j@thread.v2)'),
  messageId: z.string().describe('Message ID to react to'),
  reactionType: z
    .string()
    .describe(
      'Unicode emoji to react with (e.g., 😆 (laugh), 👍 (thumbs up), ❤️ (heart), 😮 (surprised), 😢 (sad), 😡 (angry))',
    ),
});

export const setMessageReactionTool = (server: McpServer) => {
  server.registerTool(
    'set_message_reaction',
    {
      title: 'Set Message Reaction',
      description:
        'Add a reaction (emoji) to a Teams chat message. Multiple users can react to the same message with the same or different emojis.',
      inputSchema: schema.shape,
    },
    async ({ chatId, messageId, reactionType }) => {
      const client = await graphService.getClient();
      await client.setMessageReaction({ chatId, messageId, reactionType });

      return {
        content: [
          {
            type: 'text',
            text: 'Success',
          },
        ],
      };
    },
  );
};
