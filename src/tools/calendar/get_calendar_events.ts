import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import { optimizeEvent } from './utils.js';

const schema = z.object({
  from: z
    .string()
    .optional()
    .describe(
      "Start of time range in ISO datetime UTC (e.g., '2025-01-01T00:00:00Z'). Defaults to current time if not specified.",
    ),
  to: z
    .string()
    .optional()
    .describe(
      "End of time range in ISO datetime UTC (e.g., '2025-01-31T23:59:59Z'). Defaults to 30 days from start if not specified.",
    ),
  subject: z.string().optional().describe('Filter events by subject text (partial match)'),
  limit: z.number().min(1).max(100).optional().default(50).describe('Maximum number of events to return (1-100)'),
});

export const getCalendarEventsTool = (server: McpServer) => {
  server.registerTool(
    'get_calendar_events',
    {
      title: 'Get Calendar Events',
      description:
        'List calendar events including expanded recurring meeting instances. By default, shows events from now to 30 days ahead. Recurring meetings are automatically expanded into individual occurrences. Events are returned in chronological order (earliest first). All date times are in UTC. Returns event details including subject, start, end, attendees, and location.',
      inputSchema: schema.shape,
    },
    async ({ from, to, subject, limit }) => {
      const client = await graphService.getClient();

      const events = await client.getCalendarEvents({
        from,
        to,
        subject,
        limit,
      });

      const optimizedEvents = events.map(optimizeEvent);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                totalReturned: optimizedEvents.length,
                events: optimizedEvents,
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
