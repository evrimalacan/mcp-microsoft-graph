import { homedir } from 'node:os';
import { join } from 'node:path';
import { config } from 'dotenv';

config();

export interface GraphConfig {
  clientId: string;
  tenantId: string;
  scopes: string[];
  tokenPath: string; // Legacy: No longer used for authentication (Azure Identity SDK caches tokens automatically)
}

function validateConfig(): GraphConfig {
  const clientId = process.env.MS_GRAPH_CLIENT_ID;
  const tenantId = process.env.MS_GRAPH_TENANT_ID;

  // NOTE: tokenPath is kept for backward compatibility but is no longer used for authentication
  // Azure Identity SDK (@azure/identity) automatically caches tokens in:
  //   - macOS/Linux: ~/.IdentityService/mcp-microsoft-graph
  //   - Windows: %LOCALAPPDATA%\.IdentityService\mcp-microsoft-graph
  // This cache includes both access tokens (~1 hour) and refresh tokens (~90 days)
  const tokenPath = process.env.MS_GRAPH_TOKEN_PATH || join(homedir(), '.mcp-microsoft-graph-auth.json');

  if (!clientId) {
    throw new Error('MS_GRAPH_CLIENT_ID environment variable is required');
  }

  if (!tenantId) {
    throw new Error('MS_GRAPH_TENANT_ID environment variable is required');
  }

  // Required Microsoft Graph API scopes
  // https://learn.microsoft.com/en-us/graph/permissions-reference
  const scopes = [
    'User.Read',
    'User.ReadBasic.All',
    'Team.ReadBasic.All',
    'Channel.ReadBasic.All',
    'ChannelMessage.Read.All',
    'ChannelMessage.Send',
    'TeamMember.Read.All',
    'Chat.ReadBasic',
    'Chat.ReadWrite',
    'Calendars.ReadWrite',
  ];

  return {
    clientId,
    tenantId,
    scopes,
    tokenPath,
  };
}

export const graphConfig = validateConfig();

// Re-export individual values for backward compatibility
export const CLIENT_ID = graphConfig.clientId;
export const TENANT_ID = graphConfig.tenantId;
export const SCOPES = graphConfig.scopes;
export const TOKEN_PATH = graphConfig.tokenPath;
