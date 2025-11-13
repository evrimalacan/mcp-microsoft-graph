import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createCalendarEventTool, getCalendarEventsTool } from './calendar/index.js';
import { createChatTool, searchChatsTool } from './chats/index.js';
import { listMailsTool, sendMailTool } from './mail/index.js';
import {
  getChatMessagesTool,
  searchMessagesTool,
  sendChatMessageTool,
  setMessageReactionTool,
  unsetMessageReactionTool,
} from './messages/index.js';
import { getCurrentUserTool, getUserTool, searchUsersTool } from './users/index.js';

export function registerTools(server: McpServer) {
  // User tools
  getCurrentUserTool(server);
  searchUsersTool(server);
  getUserTool(server);

  // Chat tools
  searchChatsTool(server);
  createChatTool(server);

  // Message tools
  getChatMessagesTool(server);
  sendChatMessageTool(server);
  searchMessagesTool(server);
  setMessageReactionTool(server);
  unsetMessageReactionTool(server);

  // Mail tools
  listMailsTool(server);
  sendMailTool(server);

  // Calendar tools
  getCalendarEventsTool(server);
  createCalendarEventTool(server);
}
