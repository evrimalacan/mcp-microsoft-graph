import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  contentUrls: z.array(z.string()).describe('Array of SharePoint URLs from attachment contentUrl fields'),
  targetDir: z.string().optional().describe('Directory to save files (defaults to current directory)'),
});

export const downloadSharePointFileTool = (server: McpServer) => {
  server.registerTool(
    'download_sharepoint_file',
    {
      title: 'Download SharePoint File',
      description: 'Download files attached to Teams messages via SharePoint/OneDrive.',
      inputSchema: schema.shape,
    },
    async ({ contentUrls, targetDir }) => {
      const client = await graphService.getClient();
      const dir = targetDir || process.cwd();
      await mkdir(dir, { recursive: true });

      const results = await Promise.all(
        contentUrls.map(async (contentUrl) => {
          try {
            const result = await client.downloadSharePointFile({ contentUrl });
            const filePath = join(dir, result.filename);
            await writeFile(filePath, result.data);
            return { success: true, contentUrl, path: filePath };
          } catch (error: any) {
            const message =
              error.code === 'accessDenied'
                ? 'You do not have access to this file. Please ask the owner to grant you permissions.'
                : error.message;
            return { success: false, contentUrl, error: message };
          }
        }),
      );

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(results, null, 2),
          },
        ],
      };
    },
  );
};
