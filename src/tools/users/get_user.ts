import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { OptimizedUser } from '../tools.types.js';

const schema = z.object({
  userId: z.string().describe('User ID or email address'),
});

export const getUserTool = (server: McpServer) => {
  server.registerTool(
    'get_user',
    {
      title: 'Get User',
      description:
        'Get detailed information about a specific user by their ID or email address. Returns profile information including name, email, job title, and department.',
      inputSchema: schema.shape,
    },
    async ({ userId }) => {
      const client = await graphService.getClient();
      const user = await client.getUser({ userId });

      const userSummary: OptimizedUser = {
        displayName: user.displayName,
        mail: user.mail,
        id: user.id,
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(userSummary, null, 2),
          },
        ],
      };
    },
  );
};
