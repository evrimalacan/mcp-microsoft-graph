import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  itemId: z.string().describe('The file item ID'),
  driveId: z.string().describe('The drive ID where the file is located'),
  filePath: z.string().describe('Local file path with the new content to upload'),
});

export const updateSharePointFileTool = (server: McpServer) => {
  server.registerTool(
    'update_sharepoint_file',
    {
      title: 'Update SharePoint File',
      description: 'Replace the content of an existing file.',
      inputSchema: schema.shape,
    },
    async ({ itemId, driveId, filePath }) => {
      const client = await graphService.getClient();

      try {
        const result = await client.updateSharePointFile({ itemId, driveId, filePath });
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, filename: result.filename, size: result.size }, null, 2) }],
        };
      } catch (error: any) {
        let message: string;
        if (error.code === 'accessDenied') {
          message = "Access denied. You don't have edit permissions on this file. Ask the file owner to grant you write access.";
        } else if (error.statusCode === 423 || error.code === 'resourceLocked' || error.code === 'notAllowed') {
          message = 'File is locked. It may be open in a browser or desktop app. Ask the owner to close it and try again.';
        } else {
          message = error.message || 'Unknown error';
        }
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: false, error: message }, null, 2) }],
        };
      }
    },
  );
};
