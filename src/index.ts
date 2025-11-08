#!/usr/bin/env node

import { promises as fs } from 'node:fs';
import { DeviceCodeCredential } from '@azure/identity';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CLIENT_ID, SCOPES, TENANT_ID, TOKEN_PATH } from './config.js';
import { registerTools } from './tools/index.js';

async function authenticate() {
  console.log('Microsoft Graph Authentication for MCP Server');
  console.log('='.repeat(50));

  const credential = new DeviceCodeCredential({
    clientId: CLIENT_ID,
    tenantId: TENANT_ID,
    userPromptCallback: (info) => {
      console.log('\nPlease complete authentication:');
      console.log(`Visit: ${info.verificationUri}`);
      console.log(`Enter code: ${info.userCode}`);
      console.log('\nThis prompt will automatically close once authentication is complete.');
    },
  });

  const token = await credential.getToken(SCOPES);

  if (!token) {
    throw new Error('Failed to obtain token');
  }

  const authInfo = {
    token: token.token,
    expiresAt: token.expiresOnTimestamp ? new Date(token.expiresOnTimestamp).toISOString() : undefined,
  };

  await fs.writeFile(TOKEN_PATH, JSON.stringify(authInfo, null, 2));

  console.log('\nAuthentication successful! You can now use the MCP server.');
}

async function startMcpServer() {
  const server = new McpServer({
    name: 'mcp-microsoft-graph',
    version: '1.0.0',
  });

  registerTools(server);

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Microsoft Graph MCP Server started');
}

async function main() {
  const command = process.argv[2];

  if (command === 'authenticate' || command === 'auth') {
    await authenticate();
  } else {
    await startMcpServer();
  }
}

main().catch((error) => {
  console.error('Error:', error.message);
  process.exit(1);
});
