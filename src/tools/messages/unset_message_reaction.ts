import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  chatId: z.string().describe('Chat ID (e.g. 19:meeting_Njhi..j@thread.v2)'),
  messageId: z.string().describe('Message ID to remove reaction from'),
  reactionType: z
    .string()
    .describe(
      'Unicode emoji to remove (e.g., 😆 (laugh), 👍 (thumbs up), ❤️ (heart), 😮 (surprised), 😢 (sad), 😡 (angry))',
    ),
});

export const unsetMessageReactionTool = (server: McpServer) => {
  server.registerTool(
    'unset_message_reaction',
    {
      title: 'Unset Message Reaction',
      description: 'Remove a reaction (emoji) from a Teams chat message that you previously added.',
      inputSchema: schema.shape,
    },
    async ({ chatId, messageId, reactionType }) => {
      const client = await graphService.getClient();
      await client.unsetMessageReaction({ chatId, messageId, reactionType });

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
