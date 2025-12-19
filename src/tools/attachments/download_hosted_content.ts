import { randomUUID } from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  chatId: z.string().describe('Chat ID'),
  messageId: z.string().describe('Message ID containing the images'),
  hostedContentIds: z
    .array(z.string())
    .describe('Array of hosted content IDs (from <img src> URLs, the string between /hostedContents/ and /$value)'),
  targetDir: z.string().optional().describe('Directory to save files (defaults to current directory)'),
});

export const downloadHostedContentTool = (server: McpServer) => {
  server.registerTool(
    'download_hosted_content',
    {
      title: 'Download Hosted Content',
      description:
        'Download inline images (hosted content) from a Teams message. Use this for images embedded in message body via <img> tags.',
      inputSchema: schema.shape,
    },
    async ({ chatId, messageId, hostedContentIds, targetDir }) => {
      const client = await graphService.getSDKClient();
      const dir = targetDir || '.';
      await mkdir(dir, { recursive: true });

      const results = await Promise.all(
        hostedContentIds.map(async (hostedContentId) => {
          const response = await client
            .api(`/chats/${chatId}/messages/${messageId}/hostedContents/${hostedContentId}/$value`)
            .responseType('arraybuffer' as any)
            .get();
          const data = Buffer.from(response);

          // Detect file type from magic bytes
          const isPng = data[0] === 0x89 && data[1] === 0x50;
          const isJpeg = data[0] === 0xff && data[1] === 0xd8;
          const isGif = data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46;

          let extension = '.bin';
          if (isPng) extension = '.png';
          else if (isJpeg) extension = '.jpg';
          else if (isGif) extension = '.gif';

          const filename = `${randomUUID()}${extension}`;
          const filePath = join(dir, filename);

          await writeFile(filePath, data);
          return { path: filePath, size: data.length };
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
