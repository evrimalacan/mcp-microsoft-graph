import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { Event, GraphApiResponse } from '../../types/graph.js';

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

      // Calculate date range with defaults
      const now = new Date();
      const thirtyDaysLater = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

      const startDateTime = from || now.toISOString();
      const endDateTime = to || thirtyDaysLater.toISOString();

      // calendarView requires startDateTime and endDateTime as query parameters
      let query = client
        .api('/me/calendar/calendarView')
        .query({ startDateTime, endDateTime })
        .top(limit)
        .orderby('start/dateTime');

      // Apply subject filter if specified
      if (subject) {
        query = query.filter(`contains(subject, '${subject}')`);
      }

      const response = (await query.get()) as GraphApiResponse<Event>;

      const events = (response?.value || []).map((event) => ({
        id: event.id,
        subject: event.subject,
        body: event.body?.content,
        start: event.start?.dateTime,
        end: event.end?.dateTime,
        location: event.location?.displayName,
        isOnlineMeeting: event.isOnlineMeeting,
        onlineMeetingUrl: event.onlineMeeting?.joinUrl,
        attendees: event.attendees?.map((attendee) => ({
          name: attendee.emailAddress?.name,
          email: attendee.emailAddress?.address,
          type: attendee.type,
          status: attendee.status?.response,
        })),
        organizer: {
          name: event.organizer?.emailAddress?.name,
          email: event.organizer?.emailAddress?.address,
        },
      }));

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                totalReturned: events.length,
                hasMore: !!response['@odata.nextLink'],
                events,
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
