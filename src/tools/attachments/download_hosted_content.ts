import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import { writeFile } from 'node:fs/promises';
import { join } from 'node:path';

const schema = z.object({
  chatId: z.string().describe('Chat ID'),
  messageId: z.string().describe('Message ID containing the image'),
  hostedContentId: z
    .string()
    .describe('Hosted content ID (from <img src> URL, the string between /hostedContents/ and /$value)'),
  targetDir: z.string().optional().describe('Directory to save file (defaults to current directory)'),
});

export const downloadHostedContentTool = (server: McpServer) => {
  server.registerTool(
    'download_hosted_content',
    {
      title: 'Download Hosted Content',
      description:
        'Download an inline image (hosted content) from a Teams message. Use this for images embedded in message body via <img> tags.',
      inputSchema: schema.shape,
    },
    async ({ chatId, messageId, hostedContentId, targetDir }) => {
      const client = await graphService.getSDKClient();

      // Download the content
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

      // Use short hash of hostedContentId for filename
      const shortId = hostedContentId.slice(0, 12);
      const filename = `image-${shortId}${extension}`;
      const dir = targetDir || '.';
      const filePath = join(dir, filename);

      // Save to file
      await writeFile(filePath, data);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({ path: filePath, size: data.length }, null, 2),
          },
        ],
      };
    },
  );
};
