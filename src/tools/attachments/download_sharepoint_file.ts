import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  contentUrls: z.array(z.string()).describe('SharePoint/OneDrive URLs to download'),
  targetDir: z.string().optional().describe('Directory to save files (defaults to current directory)'),
});

export const downloadSharePointFileTool = (server: McpServer) => {
  server.registerTool(
    'download_sharepoint_file',
    {
      title: 'Download SharePoint File',
      description: 'Download files from SharePoint/OneDrive URLs. Returns itemId and driveId for each file.',
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
            return {
              success: true,
              path: filePath,
              size: result.size,
              itemId: result.itemId,
              driveId: result.driveId,
            };
          } catch (error: any) {
            let message: string;
            if (error.code === 'accessDenied') {
              message = "Access denied. You don't have permission to access this file. Ask the file owner to share it with you.";
            } else if (error.code === 'itemNotFound') {
              message = 'File not found. The URL may be invalid or the file was deleted.';
            } else {
              message = error.message || 'Unknown error';
            }
            return { success: false, error: message };
          }
        }),
      );

      return {
        content: [{ type: 'text', text: JSON.stringify(results, null, 2) }],
      };
    },
  );
};
