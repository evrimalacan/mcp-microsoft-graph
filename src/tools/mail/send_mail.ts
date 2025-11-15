import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { graphService } from '../../services/graph.js';

const schema = z.object({
  to: z.array(z.string()).min(1).describe('Array of recipient email addresses'),
  subject: z.string().describe('Email subject'),
  body: z.string().describe('Email body content'),
  bodyType: z.enum(['text', 'html']).optional().default('text').describe('Body content type (text or html)'),
  cc: z.array(z.string()).optional().describe('Array of CC recipient email addresses'),
  bcc: z.array(z.string()).optional().describe('Array of BCC recipient email addresses'),
});

export const sendMailTool = (server: McpServer) => {
  server.registerTool(
    'send_mail',
    {
      title: 'Send Email',
      description:
        'Send an email message. Supports multiple recipients, CC, BCC, and both plain text and HTML formatting.',
      inputSchema: schema.shape,
    },
    async ({ to, subject, body, bodyType, cc, bcc }) => {
      const client = await graphService.getClient();

      await client.sendMail({
        to,
        subject,
        body,
        bodyType,
        cc,
        bcc,
      });

      return {
        content: [
          {
            type: 'text',
            text: `Email sent successfully to ${to.length} recipient${to.length > 1 ? 's' : ''}${
              cc ? ` (CC: ${cc.length})` : ''
            }${bcc ? ` (BCC: ${bcc.length})` : ''}`,
          },
        ],
      };
    },
  );
};
