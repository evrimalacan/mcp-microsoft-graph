import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { OptimizedEvent } from '../tools.types.js';

const schema = z.object({
  subject: z.string().describe('Event title/subject'),
  startDateTime: z.string().describe("Event start time in ISO 8601 format UTC (e.g., '2025-01-15T14:00:00Z')"),
  endDateTime: z.string().describe("Event end time in ISO 8601 format UTC (e.g., '2025-01-15T15:00:00Z')"),
  body: z.string().describe('Event description/agenda (required - no meetings without agendas)'),
  attendees: z
    .array(
      z.union([
        z.string().email().describe('Email address of attendee'),
        z
          .object({
            email: z.string().email(),
            name: z.string().optional(),
            type: z.enum(['required', 'optional', 'resource']).optional().default('required'),
          })
          .describe('Attendee object with email, name, and type'),
      ]),
    )
    .optional()
    .describe('List of attendee email addresses or attendee objects'),
  location: z.string().optional().describe("Location display name (e.g., 'Conference Room A')"),
  isOnlineMeeting: z.boolean().optional().describe('Create as Teams online meeting (auto-generates meeting link)'),
});

export const createCalendarEventTool = (server: McpServer) => {
  server.registerTool(
    'create_calendar_event',
    {
      title: 'Create Calendar Event',
      description:
        'Create a new calendar event with participants and details. All datetimes must be in UTC. Agenda (body) is required - no meetings without agendas. Can optionally create as Teams online meeting.',
      inputSchema: schema.shape,
    },
    async ({ subject, startDateTime, endDateTime, body, attendees, location, isOnlineMeeting }) => {
      const client = await graphService.getClient();

      const newEvent = await client.createCalendarEvent({
        subject,
        startDateTime,
        endDateTime,
        body,
        attendees,
        location,
        isOnlineMeeting,
      });

      const optimizedEvent: OptimizedEvent = {
        id: newEvent.id,
        subject: newEvent.subject || undefined,
        start: newEvent.start || undefined,
        end: newEvent.end || undefined,
        location: newEvent.location || undefined,
        isOnlineMeeting: newEvent.isOnlineMeeting || undefined,
        onlineMeeting: newEvent.onlineMeeting || undefined,
        attendees: newEvent.attendees || undefined,
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(optimizedEvent, null, 2),
          },
        ],
      };
    },
  );
};
