# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an MCP (Model Context Protocol) server for Microsoft Graph API integration, specifically focused on Microsoft Teams chat functionality. The server provides tools for:
- User management (search, get current user, get user details)
- Chat operations (list chats, create chats)
- Message operations (get messages, send messages, search messages)
- Mail operations (list emails, send emails)
- Calendar operations (list events, create events)

## Authentication & Token Management

### How Authentication Works

This MCP server uses **Azure Identity SDK** with **DeviceCodeCredential** for authentication. This eliminates the need for hourly re-authentication by implementing automatic token refresh.

**Key Components:**
1. **DeviceCodeCredential** - OAuth 2.0 device code flow for headless/CLI authentication
2. **Token Cache Persistence** - Tokens stored securely on disk
3. **AuthenticationRecord** - Enables silent re-authentication across restarts

**Token Lifecycle:**
- **Access Tokens**: Expire after ~1 hour, refresh automatically
- **Refresh Tokens**: Valid for ~90 days, enable silent token renewal
- **No interruptions**: Once authenticated, works for 90 days without user interaction

### First-Time Setup

Users must authenticate once:

```bash
npm run auth
# OR
npx mcp-microsoft-graph authenticate
```

This will:
1. Display a device code and URL (e.g., https://microsoft.com/devicelogin)
2. User visits URL and enters code
3. Saves **AuthenticationRecord** to `~/.mcp-microsoft-graph-auth-record.json`
4. Caches tokens in `~/.IdentityService/mcp-microsoft-graph` (macOS/Linux) or `%LOCALAPPDATA%\.IdentityService\mcp-microsoft-graph` (Windows)

### How Token Refresh Works

**Automatic Refresh Flow:**
```
1. MCP server starts → loads AuthenticationRecord
2. Tool called → needs access token
3. Check cached token:
   - If valid → use it
   - If expired → use refresh token to get new access token (SILENT)
   - If refresh token expired → prompt for device code (after ~90 days)
```

**Implementation Details:**

The `GraphService` class (src/services/graph.ts) handles everything:

```typescript
// Load AuthenticationRecord on startup
const authRecord = await loadAuthRecord();

// Create credential with persistent cache
const credential = new DeviceCodeCredential({
  clientId: CLIENT_ID,
  tenantId: TENANT_ID,
  tokenCachePersistenceOptions: {
    enabled: true,
    name: 'mcp-microsoft-graph',
  },
  authenticationRecord: authRecord, // KEY: Enables silent auth
});

// Get token (automatic refresh if expired)
const tokenResponse = await credential.getToken(SCOPES);
```

**Why This Works:**
- `@azure/identity` SDK handles all refresh logic internally
- `AuthenticationRecord` tells SDK which cached account to use
- Refresh tokens are automatically used when access tokens expire
- Only prompts for device code if refresh token is gone/expired

### Re-authentication

Users only need to re-authenticate if:
1. **90 days have passed** (refresh token expired)
2. **Credentials were revoked** (admin action)
3. **Cache files deleted** (manual cleanup)

Simply run `npm run auth` again when needed.

### Security Notes

- **Tokens encrypted**: Windows uses DPAPI, macOS uses Keychain (when available)
- **Fallback**: `unsafeAllowUnencryptedStorage: true` allows file-based cache on systems without secure storage
- **No secrets in code**: Client ID and Tenant ID are not secrets (public OAuth identifiers)
- **Refresh token rotation**: Microsoft may rotate refresh tokens periodically for security

### Troubleshooting

**"Failed to obtain token" error:**
```bash
# Solution: Re-authenticate
npm run auth
```

**MCP server not connecting:**
- Check that authentication completed successfully
- Verify `~/.mcp-microsoft-graph-auth-record.json` exists
- Check logs for specific error messages

**Token refresh failing:**
- Tokens may have been revoked by admin
- Re-authenticate: `npm run auth`

## Microsoft Graph API Tool Development

### Always Use Context7 for Documentation

**CRITICAL**: The Microsoft Graph API changes frequently. Never rely on training data knowledge alone. Always verify endpoint capabilities, query parameters, and response structures using context7.

#### Context7 Library Information
- **Library ID**: `/microsoftgraph/microsoft-graph-docs-contrib`
- **Trust Score**: 9.5/10
- **Code Snippets**: 138,345 examples
- **Content**: Official Microsoft Graph API documentation repository

### Tool Development Workflow

When creating or modifying any Microsoft Graph API tool, follow this workflow:

#### 1. Research Phase (ALWAYS REQUIRED)
Before writing or modifying any tool code:

```
Step 1: Identify the endpoint you need to work with
Step 2: Use context7 to fetch documentation for that specific endpoint
Step 3: Verify supported query parameters ($expand, $filter, $top, $orderby, etc.)
Step 4: Check response structure and available properties
Step 5: Review code examples in multiple languages (JS, Python, HTTP, etc.)
```

**Context7 Query Template**:
```typescript
// Use mcp__context7__get-library-docs with:
// - context7CompatibleLibraryID: /microsoftgraph/microsoft-graph-docs-contrib
// - topic: "[endpoint name] [specific feature]"
// - tokens: 5000-8000 (adjust based on complexity)

// Example topics:
// - "list chats endpoint $expand members query parameter"
// - "send message to chat endpoint"
// - "search messages filter parameters"
```

#### 2. Verification Phase
After researching, verify:

- ✅ Endpoint exists and is documented
- ✅ Query parameters are supported for this specific endpoint
- ✅ Response structure matches your type definitions
- ✅ Required permissions are documented
- ✅ Rate limits or special considerations are noted

#### 3. Implementation Phase
Only after verification, implement the tool following this structure:

```typescript
// src/tools/[domain]/[tool_name].ts
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { graphService } from "../../services/graph.js";
import type { OptimizedResponseType } from "../tools.types.js";

// Define schema with descriptions from official docs
// Include ISO datetime format examples for date parameters
const schema = z.object({
  paramName: z.string().describe("Description from Graph API docs"),
  dateParam: z.string().optional().describe("ISO datetime (e.g., '2025-01-01T00:00:00Z')"),
});

export const toolNameTool = (server: McpServer) => {
  server.registerTool(
    "tool_name",
    {
      title: "Human-Readable Title",
      description: "Description based on official Microsoft documentation",
      inputSchema: schema.shape,
    },
    async (params) => {
      const client = await graphService.getClient();

      // Call typed client method
      const response = await client.someMethod({
        paramName: params.paramName,
        dateParam: params.dateParam,
      });

      // Transform to optimized type for MCP response
      const optimized: OptimizedResponseType = {
        id: response.id,
        relevantField: response.relevantField,
        // Only include fields needed for MCP, reducing token usage
      };

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(optimized, null, 2),
          },
        ],
      };
    }
  );
};
```

### Common Graph API Query Parameters

Based on official documentation, these parameters are commonly available:

#### $expand
Includes related resources in the response (reduces API calls).

**Example from list_chats.ts**:
```typescript
// ✅ Verified via context7: $expand=members is valid for /me/chats
const queryParams: string[] = ["$expand=members"];
```

**Common values**:
- `members` - Include chat/team members
- `lastMessagePreview` - Include preview of last message
- `installedApps` - Include installed apps (for teams)

#### $filter
Filters results based on criteria.

**Examples**:
```typescript
// Filter by chat type
"$filter=chatType eq 'oneOnOne'"

// Filter by member name (requires $expand=members)
"$filter=members/any(o: o/displayname eq 'John Doe')"
```

#### $top
Limits the number of results returned.

**Example**:
```typescript
"$top=50" // Return only 50 most recent items
```

#### $orderby
Sorts results by specified property.

**Examples**:
```typescript
"$orderby=lastUpdatedDateTime desc"
"$orderby=createdDateTime asc"
```

### Real-World Example: list_chats Tool

This tool was verified using context7 and implements best practices:

**Location**: `src/tools/chats/list_chats.ts`

**Research Query Used**:
```
Topic: "list chats endpoint $expand members query parameter"
Library: /microsoftgraph/microsoft-graph-docs-contrib
```

**Verified Facts**:
- ✅ Endpoint: `/me/chats` or `/users/{id}/chats`
- ✅ Supports: `$expand=members`, `$filter` (contains function)
- ✅ Response includes: id, topic, chatType, members (when expanded)

**Current Implementation**:
```typescript
// Uses SDK fluent API methods
let query = client.api(`/me/chats`).expand('members');

if (searchTerm) {
  query = query.filter(`contains(topic, '${searchTerm}')`);
}

const response = await query.get();
```

**Key Features**:
- Optional `searchTerm` parameter for filtering by chat topic
- Returns empty array if no chats found
- Uses SDK fluent methods instead of manual query string construction

### Tool Maintenance Guidelines

#### When to Update Tools

1. **API Version Changes**: Microsoft Graph API releases new versions
2. **Deprecation Notices**: Endpoint or parameter deprecation warnings
3. **Feature Requests**: Users need additional query parameters
4. **Bug Reports**: Unexpected responses or errors

#### Update Process

1. **Research First**: Use context7 to check current documentation
2. **Verify Changes**: Compare with existing implementation
3. **Update Types**: Modify TypeScript types if response structure changed
4. **Update Schema**: Add/remove Zod schema parameters
5. **Test**: Verify with actual API calls
6. **Document**: Update tool descriptions and comments

### Adding New Tools

When adding a completely new tool:

1. **Identify Use Case**: What Graph API capability is needed?
2. **Research Endpoint**: Find the correct endpoint via context7
3. **Check Permissions**: Document required Microsoft Graph permissions
4. **Create Tool File**: Follow the structure in `src/tools/[domain]/`
5. **Register Tool**: Export from domain's `index.ts` barrel file
6. **Import in Main**: Add to `src/tools/index.ts` registerTools function

**Example - Adding a new tool**:

```typescript
// 1. Create: src/tools/users/update_user_status.ts
export const updateUserStatusTool = (server: McpServer) => { /* ... */ };

// 2. Export: src/tools/users/index.ts
export * from "./update_user_status.js";

// 3. Register: src/tools/index.ts
import { updateUserStatusTool } from "./users/index.js";

export function registerTools(server: McpServer) {
  // ... existing tools
  updateUserStatusTool(server);
}
```

### Microsoft Graph Permissions Reference

Each tool requires specific permissions. Document these in tool descriptions.

**Common permissions for this project**:
- `User.Read` - Read user profile
- `User.ReadBasic.All` - Read all users' basic profiles
- `Chat.ReadBasic` - Read basic chat properties
- `Chat.ReadWrite` - Read and write chats
- `ChatMessage.Read.All` - Read all chat messages
- `ChatMessage.Send` - Send chat messages

**Setting permissions**: Configure in Azure AD app registration, then update `SCOPES` in `src/config.ts`.

### Error Handling Best Practices

```typescript
// ❌ Don't catch errors unnecessarily
try {
  const client = await graphService.getClient();
  // ...
} catch (error) {
  // Over-defensive, makes code look "AI-written"
}

// ✅ Let errors bubble up naturally
const client = await graphService.getClient();
// If authentication fails, clear error message will show:
// "Not authenticated. Run: npx mcp-microsoft-graph authenticate"
```

### Testing Tools Manually

```bash
# 1. Build the project
npm run build

# 2. Ensure authenticated
npm run auth

# 3. Test via MCP client (like Claude Desktop)
# Or use the MCP inspector:
npx @modelcontextprotocol/inspector node dist/index.js
```

### Additional Resources

- **Microsoft Graph Explorer**: https://developer.microsoft.com/graph/graph-explorer
  - Interactive tool to test endpoints before implementing
- **Graph API Documentation**: https://learn.microsoft.com/graph
  - Official docs (but always verify via context7 for latest)
- **Graph SDK for JavaScript**: https://github.com/microsoftgraph/msgraph-sdk-javascript
  - We use this SDK - check for updates periodically

## Project Architecture

This project follows a **clean layered architecture** pattern inspired by mcp-bitbucket-server:

### Layer 1: Client Layer (`src/client/`)
Pure, typed API client - library-grade, framework-agnostic, reusable code.

- **`graph.client.ts`**: `GraphClient` class with typed methods for all Graph API operations
  - User operations: `getCurrentUser()`, `getUser()`, `searchUsers()`
  - Chat operations: `searchChats()`, `createChat()`
  - Message operations: `getChatMessages()`, `sendChatMessage()`, `searchMessages()`, `setMessageReaction()`, `unsetMessageReaction()`
  - Mail operations: `listMails()`, `sendMail()`
  - Calendar operations: `getCalendarEvents()`, `createCalendarEvent()`

- **`graph.types.ts`**: Complete type system
  - Re-exports from `@microsoft/microsoft-graph-types` (official SDK types)
  - API response wrapper types (`GraphApiResponse`, `SearchResponse`)
  - Method parameter types (`GetUserParams`, `SendChatMessageParams`, etc.)
  - All types prefixed by purpose (e.g., `GetUserParams` for input, `User` for output)

- **`index.ts`**: Barrel export for clean imports

**Key Principle**: The client layer has NO knowledge of MCP. It's a standalone Graph API client that could be extracted as an npm package.

### Layer 2: Service Layer (`src/services/`)
Authentication and client instantiation.

- **`graph.ts`**: `GraphService` singleton
  - `getClient()`: Returns typed `GraphClient` instance (for tools)
  - `getSDKClient()`: Returns raw SDK `Client` (for utilities that need direct API access)
  - Handles authentication, token management, and caching
  - Wraps SDK client in our typed GraphClient

### Layer 3: Tools Layer (`src/tools/`)
MCP-specific logic and response optimization.

- **`tools.types.ts`**: Optimized response types for MCP
  - `OptimizedUser`, `OptimizedChat`, `OptimizedChatMessage`, etc.
  - Strip unnecessary fields to reduce token usage
  - Separate from API types (different purpose)

- **Domain-organized tool files**:
  - `users/`: User management tools
  - `chats/`: Chat container management
  - `messages/`: Message content operations
  - `mail/`: Email operations
  - `calendar/`: Calendar event operations

**Tool Pattern**:
```typescript
import { graphService } from '../../services/graph.js';
import type { OptimizedUser } from '../tools.types.js';

export const getUserTool = (server: McpServer) => {
  server.registerTool('get_user', {}, async ({ userId }) => {
    const client = await graphService.getClient();
    const user = await client.getUser({ userId });

    // Transform to optimized type
    const optimized: OptimizedUser = {
      id: user.id,
      displayName: user.displayName,
      mail: user.mail,
    };

    return { content: [{ type: 'text', text: JSON.stringify(optimized, null, 2) }] };
  });
};
```

### Project Structure

```
src/
├── config.ts                 # Environment variables with validation
├── index.ts                  # Main entry point, server setup
├── client/                   # Layer 1: Pure API client
│   ├── graph.client.ts       # GraphClient class with typed methods
│   ├── graph.types.ts        # API types and parameter types
│   └── index.ts              # Barrel export
├── services/                 # Layer 2: Service wrapper
│   └── graph.ts              # GraphService singleton (auth + client)
├── tools/                    # Layer 3: MCP tools
│   ├── index.ts              # Main registerTools function
│   ├── tools.types.ts        # Optimized MCP response types
│   ├── users/                # User management tools
│   │   ├── index.ts
│   │   ├── get_current_user.ts
│   │   ├── search_users.ts
│   │   └── get_user.ts
│   ├── chats/                # Chat management tools
│   │   ├── index.ts
│   │   ├── search_chats.ts
│   │   └── create_chat.ts
│   ├── messages/             # Message operations tools
│   │   ├── index.ts
│   │   ├── get_chat_messages.ts
│   │   ├── send_chat_message.ts
│   │   ├── search_messages.ts
│   │   ├── set_message_reaction.ts
│   │   └── unset_message_reaction.ts
│   ├── mail/                 # Email operations tools
│   │   ├── index.ts
│   │   ├── list_mails.ts
│   │   └── send_mail.ts
│   └── calendar/             # Calendar operations tools
│       ├── index.ts
│       ├── get_calendar_events.ts
│       └── create_calendar_event.ts
└── utils/                    # Utility functions
    ├── markdown.ts           # Markdown to HTML conversion
    ├── users.ts              # User lookup utilities
    └── attachments.ts        # Attachment handling
```

### Architecture Benefits

✅ **Clean Separation**: Client layer is pure, reusable code
✅ **Type Safety**: Every API call has explicit parameter and return types
✅ **Maintainability**: API changes only affect client layer
✅ **Token Efficiency**: Tools optimize responses without polluting client types
✅ **Testability**: Client can be easily mocked
✅ **Reusability**: Client layer could be extracted as `@yourcompany/graph-client` npm package

## Current Tool Implementations

### Message Tools

#### search_messages
**Location**: `src/tools/messages/search_messages.ts`

Uses Microsoft Search API (`/search/query`) with KQL syntax. Enhanced with convenience parameters:

- `query` (optional): Raw KQL query string
- `mentions` (optional): User ID to search for mentions
- `from` (optional): ISO datetime for time range filter (e.g., '2025-01-01T00:00:00Z')
- `fromUser` (optional): Sender user ID filter
- `scope`: all/channels/chats (default: all)
- `limit`: 1-100 (default: 25)
- `enableTopResults`: boolean (default: true)

Parameters are combined with AND logic to build KQL queries.

#### get_chat_messages
**Location**: `src/tools/messages/get_chat_messages.ts`

Retrieves messages from a specific chat, sorted by creation date (newest first).

- `chatId` (required): Chat ID
- `limit`: 1-50 (default: 20)
- `from` (optional): Filter messages from this ISO datetime (e.g., '2025-01-01T00:00:00Z')
- `to` (optional): Filter messages to this ISO datetime (e.g., '2025-01-31T23:59:59Z')
- `fromUser` (optional): Filter by sender user ID

**Note**: Always uses `$orderby=createdDateTime desc` (hardcoded). Date filtering (`from`/`to`) is client-side.

#### send_chat_message
**Location**: `src/tools/messages/send_chat_message.ts`

Sends messages to a chat with support for markdown, mentions, and importance levels.

### Chat Tools

#### search_chats
**Location**: `src/tools/chats/search_chats.ts`

Searches user's chats with optional filtering by topic or member name.

- `searchTerm` (optional): Filter chats by topic name using `contains()` function
- `memberName` (optional): Filter chats by member display name (e.g., 'Martin Angelov')
- Both filters can be combined with AND logic
- Always expands members
- Returns empty array if no chats found

**Filter syntax used**:
- Topic: `contains(topic, '{searchTerm}')`
- Member: `members/any(c:contains(c/displayName, '{memberName}'))`

#### create_chat
**Location**: `src/tools/chats/create_chat.ts`

Creates new 1:1 or group chats.

### SDK Usage Best Practices

**✅ Use SDK fluent methods:**
```typescript
let query = client.api('/endpoint')
  .top(limit)
  .orderby('createdDateTime desc')
  .filter('someField eq "value"')
  .expand('members');

const response = await query.get();
```

**❌ Don't manually construct query strings:**
```typescript
// This can cause "Query option not allowed" errors
const queryString = `$top=${limit}&$orderby=createdDateTime desc`;
const response = await client.api(`/endpoint?${queryString}`).get();
```

### Parameter Naming Conventions

- Use `from`/`to` for datetime ranges (not `since`/`until`)
- Always include ISO datetime format examples in descriptions
- Use `fromUser` for sender filtering (distinguishes from datetime `from`)
- Return empty arrays instead of error messages when no results found

## Development Commands

```bash
# Development with hot reload
npm run dev

# Authenticate with Microsoft Graph
npm run auth

# Run tests (when implemented)
npm test

# Build for production
npm run build

# Lint and format (uses Biome)
npm run lint

# Type checking
npx tsc --noEmit
```

## Release Workflow

When asked to release/publish changes, follow this workflow:

1. **Check commit patterns** - Review recent commits with `git log --oneline -5` to match the commit message style (e.g., "Add X", "Fix Y", "Refactor Z")

2. **Bump version** - Use `npm version patch --no-git-tag-version` (or minor/major as appropriate)

3. **Commit changes** - Stage all files and commit with a descriptive message following the project's patterns

4. **Push to all remotes** - This project has multiple remotes:
   ```bash
   git push origin main && git push github main
   ```

5. **Tag the release** - Create and push the version tag to all remotes:
   ```bash
   git tag v1.x.x
   git push origin v1.x.x && git push github v1.x.x
   ```

6. **Publish to npm** - Request OTP from user, then:
   ```bash
   npm publish --otp=XXXXXX
   ```

**Note**: OTP (One-Time Password) will be provided by the user for npm publish.

## Testing API Endpoints Directly

To verify what data the Microsoft Graph API returns for any endpoint, use the built `graphService` to make direct API calls. This is useful for debugging and checking response structures before implementing tools.

**First, ensure the project is built:**
```bash
npm run build
```

**Then run a one-liner to test any endpoint:**

```bash
# Test listing chats with members expanded
node -e "
import('./dist/services/graph.js').then(async ({ graphService }) => {
  const client = await graphService.getSDKClient();
  const response = await client.api('/me/chats').expand('members').top(1).get();
  console.log(JSON.stringify(response.value[0], null, 2));
}).catch(console.error);
"

# Test getting a specific chat by ID
node -e "
import('./dist/services/graph.js').then(async ({ graphService }) => {
  const client = await graphService.getSDKClient();
  const response = await client.api('/me/chats/19:your-chat-id@thread.v2').expand('members').get();
  console.log(JSON.stringify(response, null, 2));
}).catch(console.error);
"

# Test any other endpoint
node -e "
import('./dist/services/graph.js').then(async ({ graphService }) => {
  const client = await graphService.getSDKClient();
  const response = await client.api('/me/messages').top(5).get();
  console.log(JSON.stringify(response.value, null, 2));
}).catch(console.error);
"
```

**Key points:**
- Uses `graphService.getSDKClient()` which returns the raw Microsoft Graph SDK client
- Authenticates automatically using stored credentials
- Supports all SDK fluent methods: `.expand()`, `.top()`, `.filter()`, `.select()`, `.orderby()`
- Output is raw JSON from the API - useful for seeing all available fields

## Environment Setup

Required environment variables in `.env.local`:

```bash
# Azure AD Application
MS_GRAPH_CLIENT_ID=your-client-id
MS_GRAPH_TENANT_ID=your-tenant-id

# Optional: Custom token storage path
MS_GRAPH_TOKEN_PATH=/path/to/custom/token.json
```

See `.env.example` for template.
