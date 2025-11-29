import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  status: z
    .enum(['Available', 'Away', 'BeRightBack', 'Busy', 'DoNotDisturb', 'Offline'])
    .describe('Status to set (Available, Busy, DoNotDisturb, Away, BeRightBack, Offline)'),
  duration: z
    .string()
    .describe('How long to keep this status (ISO 8601 format). Examples: PT1H (1 hour), PT4H (4 hours), P1D (1 day)'),
});

export const setStatusTool = (server: McpServer) => {
  server.registerTool(
    'set_status',
    {
      title: 'Set User Status',
      description: 'Set your Microsoft Teams status. Status persists even during calls and meetings.',
      inputSchema: schema.shape,
    },
    async (params) => {
      const client = await graphService.getClient();

      await client.setPreferredPresence({
        availability: params.status,
        activity: params.status,
        expirationDuration: params.duration,
      });

      return {
        content: [
          {
            type: 'text',
            text: `Status set to ${params.status} for ${params.duration}.`,
          },
        ],
      };
    },
  );
};
