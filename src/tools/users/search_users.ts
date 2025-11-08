import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { GraphApiResponse, User, UserSummary } from '../../types/graph.js';

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
      const response = (await client
        .api('/users?ConsistencyLevel=eventual')
        .search(`"mail:${query}" OR "displayName:${query}"`)
        .get()) as GraphApiResponse<User>;

      const userList: UserSummary[] = response.value!.map((user: User) => ({
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
