import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import { parseVttToText } from '../../utils/vtt.js';

const schema = z.object({
  meetingId: z.string().describe('The online meeting ID (from list_meeting_transcripts response)'),
  transcriptId: z.string().describe('The transcript ID (from list_meeting_transcripts response)'),
});

export const getMeetingTranscriptTool = (server: McpServer) => {
  server.registerTool(
    'get_meeting_transcript',
    {
      title: 'Get Meeting Transcript',
      description:
        'Get the content of a specific meeting transcript. Returns clean text with speaker names (timestamps stripped for token efficiency). Use list_meeting_transcripts first to get the meetingId and transcriptId.',
      inputSchema: schema.shape,
    },
    async ({ meetingId, transcriptId }) => {
      const client = await graphService.getClient();
      const vttContent = await client.getTranscriptContent({ meetingId, transcriptId });

      // Parse VTT to clean text format
      const cleanText = parseVttToText(vttContent);

      return {
        content: [{ type: 'text', text: cleanText }],
      };
    },
  );
};
