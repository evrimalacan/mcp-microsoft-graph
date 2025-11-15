import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { OptimizedUser } from '../tools.types.js';

const schema = z.object({
  query: z.string().describe('Search query (name or email)'),
});

export const searchUsersTool = (server: McpServer) => {
  server.registerTool(
    'search_users',
    {
      title: 'Search Users',
      description:
        'Search for users in the organization by name or email address. Returns matching users with their basic profile information.',
      inputSchema: schema.shape,
    },
    async ({ query }) => {
      const client = await graphService.getClient();
      const users = await client.searchUsers({ query });

      const userList: OptimizedUser[] = users.map((user) => ({
        displayName: user.displayName,
        mail: user.mail,
        id: user.id,
      }));

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(userList, null, 2),
          },
        ],
      };
    },
  );
};
