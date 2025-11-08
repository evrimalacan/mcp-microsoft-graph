# MCP Microsoft Graph

An MCP (Model Context Protocol) server for Microsoft Graph API integration, focused on Microsoft Teams functionality.

## Features

This MCP server provides tools for:

- **User Management**: Search users, get user details, get current user profile
- **Chat Operations**: Search and create chats (1:1 and group conversations)
- **Message Operations**: Get messages, send messages, search across Teams messages

## Installation

```bash
npm install
npm run build
```

## Authentication

Before using the server, you need to authenticate with Microsoft Graph:

```bash
npm run auth
```

This will open a browser window for you to sign in with your Microsoft account.

## Configuration

Create a `.env.local` file with your Azure AD application credentials:

```bash
# Azure AD Application
MS_GRAPH_CLIENT_ID=your-client-id
MS_GRAPH_TENANT_ID=your-tenant-id

# Optional: Custom token storage path
MS_GRAPH_TOKEN_PATH=/path/to/custom/token.json
```

See `.env.example` for a template.

### Required Permissions

Configure these permissions in your Azure AD app registration:

- `User.Read` - Read user profile
- `User.ReadBasic.All` - Read all users' basic profiles
- `Chat.ReadBasic` - Read basic chat properties
- `Chat.ReadWrite` - Read and write chats
- `ChatMessage.Read.All` - Read all chat messages
- `ChatMessage.Send` - Send chat messages

## Usage

### Development Mode

```bash
npm run dev
```

### Production

```bash
npm run build
node dist/index.js
```

### Testing with MCP Inspector

```bash
npx @modelcontextprotocol/inspector node dist/index.js
```

## Available Tools

### User Tools

#### get_current_user
Get the authenticated user's profile information.

#### search_users
Search for users by name or email address.

#### get_user
Get detailed information about a specific user.

### Chat Tools

#### search_chats
Search for chats with optional filtering:
- `searchTerm`: Filter by chat topic name
- `memberName`: Filter by member display name
- Both filters can be combined

#### create_chat
Create new 1:1 or group chats.

### Message Tools

#### get_chat_messages
Retrieve messages from a specific chat:
- `chatId` (required): The chat ID
- `limit`: Max number of messages (default: 20)
- `from`/`to`: ISO datetime filters
- `fromUser`: Filter by sender

#### send_chat_message
Send messages to a chat with support for:
- Markdown formatting
- @mentions
- Importance levels (normal, high, urgent)

#### search_messages
Search across all Teams messages using KQL syntax:
- `query`: Raw KQL query string
- `mentions`: Search for mentions of a user
- `from`: ISO datetime filter
- `fromUser`: Filter by sender
- `scope`: all/channels/chats
- `limit`: Max results (default: 25)

## Project Structure

```
src/
├── config.ts                 # Environment variables with validation
├── index.ts                  # Main entry point, server setup
├── services/
│   └── graph.ts             # GraphService singleton for API client
├── tools/
│   ├── index.ts             # Main registerTools function
│   ├── users/               # User management tools
│   ├── chats/               # Chat management tools
│   └── messages/            # Message operations tools
└── types/
    └── graph.ts             # TypeScript types for Graph API responses
```

## Development Commands

```bash
# Development with hot reload
npm run dev

# Authenticate with Microsoft Graph
npm run auth

# Run tests
npm test

# Build for production
npm run build

# Type checking
npx tsc --noEmit
```

## Resources

- [Microsoft Graph Explorer](https://developer.microsoft.com/graph/graph-explorer) - Interactive API testing tool
- [Microsoft Graph Documentation](https://learn.microsoft.com/graph) - Official API documentation
- [Graph SDK for JavaScript](https://github.com/microsoftgraph/msgraph-sdk-javascript) - Official SDK

## License

See [LICENSE](LICENSE) file for details.
