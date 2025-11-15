import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';
import type { OptimizedMail } from '../tools.types.js';

const schema = z.object({
  limit: z.number().min(1).max(50).optional().default(20).describe('Number of messages to retrieve'),
  since: z.string().optional().describe("Filter messages from this ISO datetime (e.g., '2025-01-01T00:00:00Z')"),
  until: z.string().optional().describe("Filter messages to this ISO datetime (e.g., '2025-01-31T23:59:59Z')"),
  unreadOnly: z.boolean().optional().describe('Show only unread messages'),
  hasAttachments: z.boolean().optional().describe('Filter messages with attachments'),
  from: z.string().optional().describe('Filter by sender email address'),
  includeBody: z
    .boolean()
    .optional()
    .default(false)
    .describe(
      'Include full email body content in response. By default, only bodyPreview (first ~255 characters) is returned to reduce response size. Set to true to get complete email content including links and full text.',
    ),
  bodyFormat: z
    .enum(['text', 'html'])
    .optional()
    .default('html')
    .describe(
      'Format of the email body content. "html" returns formatted HTML content (default), "text" returns plain text without formatting. Only applies when includeBody is true.',
    ),
});

export const listMailsTool = (server: McpServer) => {
  server.registerTool(
    'list_mails',
    {
      title: 'List Mail Messages',
      description:
        "Retrieve mail messages from the user's mailbox. Returns message metadata including subject, sender, received date, body content (preview or full), read status, and attachment status. By default returns bodyPreview (truncated to ~255 chars). Use includeBody parameter to get complete email content.",
      inputSchema: schema.shape,
    },
    async ({ limit, since, until, unreadOnly, hasAttachments, from, includeBody, bodyFormat }) => {
      const client = await graphService.getClient();

      const response = await client.listMails({
        limit,
        since,
        until,
        unreadOnly,
        hasAttachments,
        from,
        includeBody,
        bodyFormat,
      });

      const messages: OptimizedMail[] = (response?.value || []).map((message: any) => {
        const mappedMessage: OptimizedMail = {
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
        };

        // Include full body content if it was requested and is present
        if (includeBody && message.body) {
          mappedMessage.body = {
            contentType: message.body.contentType, // 'text' or 'html'
            content: message.body.content, // Full email body content
          };
        }

        return mappedMessage;
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(messages, null, 2),
          },
        ],
      };
    },
  );
};
