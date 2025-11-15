import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { SearchHit } from '../../client/graph.types.js';
import { graphService } from '../../services/graph.js';
import type { OptimizedSearchHit, OptimizedSearchResult } from '../tools.types.js';

const schema = z.object({
  query: z
    .string()
    .optional()
    .describe(
      "Search query. Supports KQL syntax like 'from:user mentions:userId hasAttachment:true'. Optional if using filter parameters.",
    ),
  scope: z.enum(['all', 'channels', 'chats']).optional().default('all').describe('Scope of search'),
  limit: z.number().min(1).max(100).optional().default(25).describe('Number of results to return'),
  enableTopResults: z.boolean().optional().default(true).describe('Enable relevance-based ranking'),
  mentions: z.string().optional().describe('User ID to search for mentions of that user'),
  from: z.string().optional().describe("Filter messages from this ISO datetime onwards (e.g., '2025-01-01T00:00:00Z')"),
  fromUser: z.string().optional().describe('Sender user ID filter'),
});

export const searchMessagesTool = (server: McpServer) => {
  server.registerTool(
    'search_messages',
    {
      title: 'Search Messages',
      description:
        'Search for messages across all Microsoft Teams channels and chats using Microsoft Search API. Supports advanced KQL syntax and filter parameters for mentions, time range, and sender filtering.',
      inputSchema: schema.shape,
    },
    async ({ query, scope, limit, enableTopResults, mentions, from, fromUser }) => {
      const client = await graphService.getClient();

      const response = await client.searchMessages({
        query,
        scope,
        limit,
        enableTopResults,
        mentions,
        from,
        fromUser,
      });

      if (!response?.value?.length || !response.value[0]?.hitsContainers?.length) {
        return {
          content: [
            {
              type: 'text',
              text: 'No messages found matching your search criteria.',
            },
          ],
        };
      }

      const hits = response.value[0].hitsContainers[0].hits;
      const searchResults: OptimizedSearchHit[] = hits.map((hit: SearchHit) => ({
        id: hit.resource.id,
        rank: hit.rank,
        content: hit.summary,
        from: hit.resource.from?.emailAddress?.address,
        createdDateTime: hit.resource.createdDateTime,
        chatId: hit.resource.chatId,
        teamId: hit.resource.channelIdentity?.teamId,
        channelId: hit.resource.channelIdentity?.channelId,
      }));

      const result: OptimizedSearchResult = {
        query,
        scope: scope || 'all',
        totalResults: response.value[0].hitsContainers[0].total,
        results: searchResults,
        moreResultsAvailable: response.value[0].hitsContainers[0].moreResultsAvailable,
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    },
  );
};
