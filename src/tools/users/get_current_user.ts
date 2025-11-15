import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { OptimizedUser } from '../tools.types.js';

const schema = z.object({});

export const getCurrentUserTool = (server: McpServer) => {
  server.registerTool(
    'get_current_user',
    {
      title: 'Get Current User',
      description:
        "Get the current authenticated user's profile information including display name, email, job title, and department.",
      inputSchema: schema.shape,
    },
    async () => {
      const client = await graphService.getClient();
      const user = await client.getCurrentUser();

      const userSummary: OptimizedUser = {
        displayName: user.displayName,
        mail: user.mail,
        id: user.id,
        jobTitle: user.jobTitle,
        department: user.department,
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
