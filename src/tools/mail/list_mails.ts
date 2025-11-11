import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  limit: z.number().min(1).max(50).optional().default(20).describe('Number of messages to retrieve'),
  since: z.string().optional().describe("Filter messages from this ISO datetime (e.g., '2025-01-01T00:00:00Z')"),
  until: z.string().optional().describe("Filter messages to this ISO datetime (e.g., '2025-01-31T23:59:59Z')"),
  unreadOnly: z.boolean().optional().describe('Show only unread messages'),
  hasAttachments: z.boolean().optional().describe('Filter messages with attachments'),
  from: z.string().optional().describe('Filter by sender email address'),
});

export const listMailsTool = (server: McpServer) => {
  server.registerTool(
    'list_mails',
    {
      title: 'List Mail Messages',
      description:
        "Retrieve mail messages from the user's mailbox. Returns message metadata including subject, sender, received date, body preview, read status, and attachment status.",
      inputSchema: schema.shape,
    },
    async ({ limit, since, until, unreadOnly, hasAttachments, from }) => {
      const client = await graphService.getClient();

      let query = client
        .api('/me/messages')
        .top(limit)
        .orderby('receivedDateTime desc')
        .select('id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments');

      // Build filter conditions following Microsoft Graph API restrictions
      // Reference: https://devblogs.microsoft.com/microsoft365dev/update-to-filtering-and-sorting-rest-api/
      // Reference: https://learn.microsoft.com/en-us/graph/api/user-list-messages
      //
      // IMPORTANT RULES when combining $filter and $orderby:
      // 1. All fields in $orderby must also appear in $filter
      // 2. Fields in $orderby must appear first in $filter, in the same order
      // 3. Fields exclusive to $filter must come after those in $orderby
      //
      // Violating these rules causes "InefficientFilter" error:
      // "The restriction or sort order is too complex for this operation"
      const filters: string[] = [];

      // Always include receivedDateTime in filter since it's in $orderby
      // Without this, filtering by sender alone would fail with InefficientFilter error
      if (since) {
        filters.push(`receivedDateTime ge ${since}`);
      } else {
        // Use a broad date range if not specified to satisfy API requirement
        filters.push(`receivedDateTime ge 1900-01-01T00:00:00Z`);
      }

      if (until) {
        filters.push(`receivedDateTime le ${until}`);
      }

      // Other filters must come after receivedDateTime (the orderby field)
      if (from) {
        filters.push(`from/emailAddress/address eq '${from}'`);
      }

      if (unreadOnly) {
        filters.push('isRead eq false');
      }

      if (hasAttachments !== undefined) {
        filters.push(`hasAttachments eq ${hasAttachments}`);
      }

      // Always apply filter (we always have at least receivedDateTime)
      query = query.filter(filters.join(' and '));

      const response = await query.get();

      const messages = (response?.value || []).map((message: any) => ({
        id: message.id,
        subject: message.subject,
        from: {
          name: message.from?.emailAddress?.name,
          address: message.from?.emailAddress?.address,
        },
        receivedDateTime: message.receivedDateTime,
        bodyPreview: message.bodyPreview,
        isRead: message.isRead,
        hasAttachments: message.hasAttachments,
      }));

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(response, null, 2),
          },
        ],
      };
    },
  );
};
