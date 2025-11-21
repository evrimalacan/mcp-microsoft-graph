import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  chatId: z.string().describe('The meeting chat ID (e.g., "19:meeting_YzAxMGQ3NDct...@thread.v2")'),
});

export const listMeetingTranscriptsTool = (server: McpServer) => {
  server.registerTool(
    'list_meeting_transcripts',
    {
      title: 'List Meeting Transcripts',
      description:
        'List all transcripts for a meeting chat. Returns the meeting ID and a list of transcripts with their IDs and timestamps. Use this to find available transcripts before fetching content.',
      inputSchema: schema.shape,
    },
    async ({ chatId }) => {
      const client = await graphService.getClient();

      // Step 1: Get the chat to find joinWebUrl
      const chat = await client.getChat({ chatId });
      const joinWebUrl = chat.onlineMeetingInfo?.joinWebUrl;

      if (!joinWebUrl) {
        return {
          content: [{ type: 'text', text: JSON.stringify({ error: 'Chat is not a meeting or has no joinWebUrl' }) }],
        };
      }

      // Step 2: Find the online meeting by joinWebUrl
      const meeting = await client.getMeetingByJoinUrl({ joinWebUrl });

      if (!meeting || !meeting.id) {
        return {
          content: [{ type: 'text', text: JSON.stringify({ error: 'Online meeting not found for this chat' }) }],
        };
      }

      // Step 3: Get transcripts for this meeting
      const transcripts = await client.listTranscripts({ meetingId: meeting.id });

      const result = {
        meetingId: meeting.id,
        transcripts: transcripts.map((t) => ({
          id: t.id,
          createdDateTime: t.createdDateTime,
          // SDK type is incomplete, but API returns endDateTime
          endDateTime: (t as unknown as { endDateTime?: string }).endDateTime,
        })),
      };

      return {
        content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
      };
    },
  );
};
