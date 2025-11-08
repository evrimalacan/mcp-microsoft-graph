import { promises as fs } from 'node:fs';
import { Client } from '@microsoft/microsoft-graph-client';
import { TOKEN_PATH } from '../config.js';

interface StoredAuthInfo {
  token: string;
  expiresAt?: string;
}

export class GraphService {
  private client: Client | null = null;

  async getClient(): Promise<Client> {
    if (this.client) {
      return this.client;
    }

    const authData = await fs.readFile(TOKEN_PATH, 'utf8');
    const authInfo: StoredAuthInfo = JSON.parse(authData);

    if (authInfo.expiresAt && new Date(authInfo.expiresAt) <= new Date()) {
      throw new Error('Token expired. Run: npx mcp-microsoft-graph authenticate');
    }

    if (!authInfo.token) {
      throw new Error('Not authenticated. Run: npx mcp-microsoft-graph authenticate');
    }

    this.client = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => authInfo.token,
      },
    });

    return this.client;
  }
}

export const graphService = new GraphService();
