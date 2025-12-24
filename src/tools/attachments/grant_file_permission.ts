import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  itemId: z.string().describe('The file item ID'),
  driveId: z.string().describe('The drive ID where the file is located'),
  emails: z.array(z.string()).describe('Email addresses of users to grant access to'),
  role: z.enum(['read', 'write']).describe('Permission level: read (view-only) or write (can edit)'),
});

export const grantFilePermissionTool = (server: McpServer) => {
  server.registerTool(
    'grant_file_permission',
    {
      title: 'Grant File Permission',
      description: 'Grant read or write permissions on a file to one or more users.',
      inputSchema: schema.shape,
    },
    async ({ itemId, driveId, emails, role }) => {
      const client = await graphService.getClient();

      try {
        await client.grantFilePermission({ itemId, driveId, emails, role });
        const roleText = role === 'write' ? 'Edit' : 'View';
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true, message: `${roleText} permission granted to ${emails.join(', ')}` }, null, 2) }],
        };
      } catch (error: any) {
        let message: string;
        if (error.code === 'accessDenied') {
          message = 'Access denied. You may not have permission to share this file.';
        } else if (error.code === 'itemNotFound') {
          message = 'File not found. The itemId or driveId may be invalid.';
        } else if (error.code === 'invalidRequest' && error.message?.includes('could not be resolved')) {
          message = 'One or more users not found. Check the email addresses.';
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
