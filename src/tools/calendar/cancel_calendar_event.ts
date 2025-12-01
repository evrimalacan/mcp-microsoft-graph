import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  eventId: z.string().describe('The ID of the calendar event to cancel'),
  comment: z.string().optional().describe('Optional message to send to attendees explaining the cancellation'),
});

export const cancelCalendarEventTool = (server: McpServer) => {
  server.registerTool(
    'cancel_calendar_event',
    {
      title: 'Cancel Calendar Event',
      description: 'Cancel a calendar event and send a cancellation message to all attendees.',
      inputSchema: schema.shape,
    },
    async ({ eventId, comment }) => {
      const client = await graphService.getClient();

      await client.cancelCalendarEvent({ eventId, comment });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: true,
                message: 'Event cancelled successfully.',
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
