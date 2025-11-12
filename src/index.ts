#!/usr/bin/env node

import { homedir } from 'node:os';
import { join } from 'node:path';
import { DeviceCodeCredential, useIdentityPlugin } from '@azure/identity';
import { cachePersistencePlugin } from '@azure/identity-cache-persistence';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CLIENT_ID, SCOPES, TENANT_ID } from './config.js';
import { registerTools } from './tools/index.js';

// Register the persistent cache plugin (must be done once at module load)
useIdentityPlugin(cachePersistencePlugin);

/**
 * Get the token cache location based on the operating system.
 */
function getTokenCachePath(): string {
  const platform = process.platform;

  if (platform === 'win32') {
    return join(process.env.LOCALAPPDATA || join(homedir(), 'AppData', 'Local'), '.IdentityService');
  }

  // macOS and Linux
  return join(homedir(), '.IdentityService');
}

/**
 * Authenticate with Microsoft Graph using device code flow.
 *
 * This function:
 * 1. Prompts user to visit a URL and enter a device code
 * 2. Obtains access token and refresh token
 * 3. Saves AuthenticationRecord for silent re-authentication
 * 4. Tokens are automatically cached by Azure Identity SDK in ~/.IdentityService/
 * 5. Tokens will auto-refresh for ~90 days without re-authentication
 */
async function authenticate() {
  console.log('Microsoft Graph Authentication for MCP Server');
  console.log('='.repeat(50));

  const credential = new DeviceCodeCredential({
    clientId: CLIENT_ID,
    tenantId: TENANT_ID,
    tokenCachePersistenceOptions: {
      enabled: true,
      name: 'mcp-microsoft-graph',  // Named cache for this application
      unsafeAllowUnencryptedStorage: true,  // Allow unencrypted storage if keychain unavailable
    },
    userPromptCallback: (info) => {
      console.log('\nPlease complete authentication:');
      console.log(`Visit: ${info.verificationUri}`);
      console.log(`Enter code: ${info.userCode}`);
      console.log('\nThis prompt will automatically close once authentication is complete.');
    },
  });

  // Authenticate and get AuthenticationRecord
  // This is the KEY step that enables silent re-authentication!
  const authRecord = await credential.authenticate(SCOPES);

  if (!authRecord) {
    throw new Error('Failed to authenticate');
  }

  // Save the AuthenticationRecord to file for future silent authentication
  const { promises: fs } = await import('node:fs');
  const { serializeAuthenticationRecord } = await import('@azure/identity');
  const authRecordPath = join(homedir(), '.mcp-microsoft-graph-auth-record.json');
  const serialized = serializeAuthenticationRecord(authRecord);
  await fs.writeFile(authRecordPath, serialized, 'utf-8');

  // Display success message with cache information
  const cachePath = getTokenCachePath();
  console.log('\n✅ Authentication successful!');
  console.log(`\n👤 Authenticated as: ${authRecord.username}`);
  console.log('\n📦 Token Cache Information:');
  console.log(`   Token cache: ${cachePath}/mcp-microsoft-graph`);
  console.log(`   Auth record: ${authRecordPath}`);
  console.log('\n🔄 Automatic Token Refresh:');
  console.log('   - Access tokens refresh automatically every ~1 hour');
  console.log('   - Refresh tokens valid for ~90 days');
  console.log('   - Silent re-authentication (no device code needed!)');
  console.log('\n🎯 Next Steps:');
  console.log('   1. Start using the MCP server (no more hourly re-auth!)');
  console.log('   2. Only re-authenticate after ~90 days or if credentials are revoked');
  console.log('\n💡 Tip: If you see authentication errors, just run this command again.');
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
