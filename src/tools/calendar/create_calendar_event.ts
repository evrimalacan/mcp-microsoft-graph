import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { Attendee, Event } from '../../types/graph.js';

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

      // Normalize attendees from email strings to proper attendee objects
      const normalizedAttendees: Attendee[] | undefined = attendees?.map((attendee) => {
        if (typeof attendee === 'string') {
          // Simple email string
          return {
            emailAddress: {
              address: attendee,
            },
            type: 'required',
          } as Attendee;
        }
        // Attendee object
        return {
          emailAddress: {
            address: attendee.email,
            name: attendee.name,
          },
          type: attendee.type || 'required',
        } as Attendee;
      });

      // Build event payload
      const eventPayload: Partial<Event> = {
        subject,
        body: {
          contentType: 'text',
          content: body,
        },
        start: {
          dateTime: startDateTime,
          timeZone: 'UTC',
        },
        end: {
          dateTime: endDateTime,
          timeZone: 'UTC',
        },
        attendees: normalizedAttendees,
      };

      if (location) {
        eventPayload.location = {
          displayName: location,
        };
      }

      if (isOnlineMeeting !== undefined) {
        eventPayload.isOnlineMeeting = isOnlineMeeting;
      }

      const newEvent = (await client.api('/me/events').post(eventPayload)) as Event;

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                id: newEvent.id,
                subject: newEvent.subject,
                body: newEvent.body?.content,
                start: newEvent.start?.dateTime,
                end: newEvent.end?.dateTime,
                location: newEvent.location?.displayName,
                isOnlineMeeting: newEvent.isOnlineMeeting,
                onlineMeetingUrl: newEvent.onlineMeeting?.joinUrl,
                attendees: newEvent.attendees?.map((attendee) => ({
                  name: attendee.emailAddress?.name,
                  email: attendee.emailAddress?.address,
                  type: attendee.type,
                })),
                webLink: newEvent.webLink,
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
