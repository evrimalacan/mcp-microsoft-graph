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
        'Upload a new file to OneDrive and create a sharing link. Use this when creating/generating files to share with users.',
      inputSchema: schema.shape,
    },
    async ({ filePath }) => {
      const client = await graphService.getClient();
      const result = await client.uploadToSharePoint({ filePath });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: true,
                itemId: result.itemId,
                driveId: result.driveId,
                shareUrl: result.shareUrl,
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
