import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  filePath: z.string().describe('Local file path to upload'),
});

export const uploadToSharePointTool = (server: McpServer) => {
  server.registerTool(
    'upload_to_sharepoint',
    {
      title: 'Upload to SharePoint',
      description:
        'Upload a file to OneDrive/SharePoint and create a sharing link. Returns file info for use with send_chat_message attachments.',
      inputSchema: schema.shape,
    },
    async ({ filePath }) => {
      const client = await graphService.getClient();

      const result = await client.uploadToSharePoint({ filePath });

      return {
        content: [
          {
            type: 'text',
            text: result.contentUrl,
          },
        ],
      };
    },
  );
};
