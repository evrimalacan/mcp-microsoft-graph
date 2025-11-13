import { promises as fs } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import {
  type AuthenticationRecord,
  DeviceCodeCredential,
  deserializeAuthenticationRecord,
  useIdentityPlugin,
} from '@azure/identity';
import { cachePersistencePlugin } from '@azure/identity-cache-persistence';
import { Client } from '@microsoft/microsoft-graph-client';
import { CLIENT_ID, SCOPES, TENANT_ID } from '../config.js';

// Register the persistent cache plugin (must be done once at module load)
useIdentityPlugin(cachePersistencePlugin);

/**
 * GraphService provides Microsoft Graph API client with automatic token refresh.
 *
 * Uses DeviceCodeCredential with persistent token caching AND AuthenticationRecord:
 * - Tokens cached in ~/.IdentityService/mcp-microsoft-graph (macOS/Linux) or %LOCALAPPDATA%\.IdentityService\ (Windows)
 * - AuthenticationRecord cached in ~/.mcp-microsoft-graph-auth-record.json
 * - Access tokens refresh automatically when they expire (~1 hour)
 * - Refresh tokens valid for ~90 days
 * - Silent re-authentication (no device code prompt!)
 *
 * This eliminates the need for hourly re-authentication!
 */
export class GraphService {
  private client: Client | null = null;
  private credential: DeviceCodeCredential | null = null;

  /**
   * Load the AuthenticationRecord from disk if it exists.
   * This enables silent authentication without device code prompts.
   */
  private async loadAuthRecord(): Promise<AuthenticationRecord | null> {
    try {
      const authRecordPath = join(homedir(), '.mcp-microsoft-graph-auth-record.json');
      const content = await fs.readFile(authRecordPath, 'utf-8');
      return deserializeAuthenticationRecord(content);
    } catch (_error) {
      // Auth record doesn't exist or can't be read - user needs to authenticate
      return null;
    }
  }

  async getClient(): Promise<Client> {
    if (this.client) {
      return this.client;
    }

    // Load the AuthenticationRecord if it exists
    const authRecord = await this.loadAuthRecord();

    // Create DeviceCodeCredential instance with persistent cache enabled
    // If we have an AuthenticationRecord, pass it for silent authentication
    this.credential = new DeviceCodeCredential({
      clientId: CLIENT_ID,
      tenantId: TENANT_ID,
      tokenCachePersistenceOptions: {
        enabled: true,
        name: 'mcp-microsoft-graph', // Named cache for this application
        unsafeAllowUnencryptedStorage: true, // Allow unencrypted storage if keychain unavailable
      },
      // KEY: Pass the AuthenticationRecord to enable silent authentication!
      authenticationRecord: authRecord || undefined,
    });

    // Create Graph client with auth provider that uses credential
    this.client = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          try {
            // credential.getToken() automatically:
            // 1. Returns cached access token if still valid
            // 2. Refreshes using cached refresh token if access token expired
            // 3. Only prompts for device code if refresh token expired (after ~90 days)
            const tokenResponse = await this.credential?.getToken(SCOPES);

            if (!tokenResponse) {
              throw new Error('Failed to obtain token. Run: npx mcp-microsoft-graph authenticate');
            }

            return tokenResponse.token;
          } catch (_error) {
            // If token acquisition fails, provide helpful error message
            throw new Error(
              'Authentication failed. This could mean:\n' +
                "  1. You haven't authenticated yet\n" +
                '  2. Your refresh token expired (after ~90 days)\n' +
                '  3. Your credentials were revoked\n\n' +
                'Solution: Run "npx mcp-microsoft-graph authenticate"',
            );
          }
        },
      },
    });

    return this.client;
  }

  /**
   * Clear the cached client (forces recreation on next getClient call).
   * Useful for testing or when you want to force token refresh.
   */
  clearCache(): void {
    this.client = null;
    this.credential = null;
  }
}

export const graphService = new GraphService();
