import { writeFile, mkdir } from 'node:fs/promises';
import { join } from 'node:path';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  contentUrl: z.string().describe('SharePoint URL from the attachment contentUrl field'),
  targetDir: z.string().optional().describe('Directory to save file (defaults to current directory)'),
});

export const downloadSharePointFileTool = (server: McpServer) => {
  server.registerTool(
    'download_sharepoint_file',
    {
      title: 'Download SharePoint File',
      description: 'Download a file attached to a Teams message via SharePoint/OneDrive.',
      inputSchema: schema.shape,
    },
    async ({ contentUrl, targetDir }) => {
      const client = await graphService.getClient();

      const result = await client.downloadSharePointFile({ contentUrl });

      // Determine target directory
      const dir = targetDir || process.cwd();
      await mkdir(dir, { recursive: true });

      // Save file
      const filePath = join(dir, result.filename);
      await writeFile(filePath, result.data);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                path: filePath,
                filename: result.filename,
                size: result.size,
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
