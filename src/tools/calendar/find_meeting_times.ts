import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { OptimizedMeetingTimeSuggestion } from '../tools.types.js';

const attendeeSchema = z.object({
  email: z.string().email().describe('Email address of the attendee'),
  type: z.enum(['required', 'optional']).describe('Whether the attendee is required or optional'),
});

const schema = z.object({
  attendees: z
    .array(attendeeSchema)
    .min(1)
    .describe(
      'Array of attendees with their email addresses and type (required or optional). Example: [{ email: "user@domain.com", type: "required" }]',
    ),
  startDateTime: z
    .string()
    .describe("Start of time window to search within (ISO 8601 format in UTC, e.g., '2025-11-28T07:00:00Z')"),
  endDateTime: z
    .string()
    .describe("End of time window to search within (ISO 8601 format in UTC, e.g., '2025-11-29T00:00:00Z')"),
  meetingDuration: z
    .string()
    .describe(
      "Meeting duration in ISO 8601 format. Examples: 'PT30M' for 30 minutes, 'PT1H' for 1 hour, 'PT1H30M' for 1.5 hours",
    ),
});

export const findMeetingTimesTool = (server: McpServer) => {
  server.registerTool(
    'find_meeting_times',
    {
      title: 'Find Available Meeting Times',
      description:
        'Find optimal meeting times for attendees by analyzing their calendars. Returns up to 10 time slot suggestions with confidence scores. Only searches during working hours. All required attendees must be available (100% threshold). The authenticated user (organizer) does not affect availability.',
      inputSchema: schema.shape,
    },
    async ({ attendees, startDateTime, endDateTime, meetingDuration }) => {
      const client = await graphService.getClient();

      const response = await client.findMeetingTimes({
        attendees,
        startDateTime,
        endDateTime,
        meetingDuration,
      });

      // Optimize response: strip unnecessary fields to reduce token usage
      const optimizedSuggestions: OptimizedMeetingTimeSuggestion[] = (response.meetingTimeSuggestions || []).map(
        (suggestion) => ({
          confidence: suggestion.confidence || 0,
          start: suggestion.meetingTimeSlot?.start?.dateTime || '',
          end: suggestion.meetingTimeSlot?.end?.dateTime || '',
        }),
      );

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(optimizedSuggestions, null, 2),
          },
        ],
      };
    },
  );
};
