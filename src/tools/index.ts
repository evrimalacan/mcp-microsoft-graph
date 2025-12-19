import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { downloadHostedContentTool, downloadSharePointFileTool, uploadToSharePointTool } from './attachments/index.js';
import {
  cancelCalendarEventTool,
  createCalendarEventTool,
  findMeetingTimesTool,
  getCalendarEventsTool,
} from './calendar/index.js';
import { createChatTool, getChatTool, searchChatsTool } from './chats/index.js';
import { listMailsTool, sendMailTool } from './mail/index.js';
import {
  getChatMessagesTool,
  getMessageTool,
  searchMessagesTool,
  sendChatMessageTool,
  setMessageReactionTool,
  unsetMessageReactionTool,
  updateChatMessageTool,
} from './messages/index.js';
import { getMeetingTranscriptTool, listMeetingTranscriptsTool } from './transcripts/index.js';
import { getCurrentUserTool, getUserTool, searchUsersTool, setStatusTool } from './users/index.js';

export function registerTools(server: McpServer) {
  // User tools
  getCurrentUserTool(server);
  searchUsersTool(server);
  getUserTool(server);
  setStatusTool(server);

  // Chat tools
  getChatTool(server);
  searchChatsTool(server);
  createChatTool(server);

  // Message tools
  getChatMessagesTool(server);
  getMessageTool(server);
  sendChatMessageTool(server);
  updateChatMessageTool(server);
  searchMessagesTool(server);
  setMessageReactionTool(server);
  unsetMessageReactionTool(server);

  // Mail tools
  listMailsTool(server);
  sendMailTool(server);

  // Calendar tools
  getCalendarEventsTool(server);
  createCalendarEventTool(server);
  cancelCalendarEventTool(server);
  findMeetingTimesTool(server);

  // Transcript tools
  listMeetingTranscriptsTool(server);
  getMeetingTranscriptTool(server);

  // Attachment tools
  downloadSharePointFileTool(server);
  downloadHostedContentTool(server);
  uploadToSharePointTool(server);
}
