# Microsoft Graph MCP

A **Model Context Protocol (MCP)** server that connects AI assistants to Microsoft Graph API. Search Teams messages, manage chats, send messages, and interact with your Microsoft 365 environment through natural language in Claude and other AI assistants.

## 🚀 Quick Start

### Configuration

#### Claude Code

Add the server using the Claude Code CLI:

```bash
claude mcp add -s user \
    microsoft-graph \
    npx mcp-microsoft-graph@latest \
    -e "MS_GRAPH_CLIENT_ID=your_client_id" \
    -e "MS_GRAPH_TENANT_ID=your_tenant_id"
```

#### Manual Configuration (Any MCP Client)

Alternatively, add this configuration to your MCP client's configuration file:

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "npx",
      "args": ["mcp-microsoft-graph@latest"],
      "type": "stdio",
      "env": {
        "MS_GRAPH_CLIENT_ID": "your_client_id",
        "MS_GRAPH_TENANT_ID": "your_tenant_id"
      }
    }
  }
}
```

### Get Your Azure AD Credentials

To get your `CLIENT_ID` and `TENANT_ID`:

1. Go to [Azure Portal](https://portal.azure.com/) → **Azure Active Directory** → **App registrations**
2. Click **New registration**
3. Give it a name (e.g., "MCP Microsoft Graph")
4. Set **Supported account types** to "Single tenant"
5. Click **Register**
6. Copy the **Application (client) ID** - this is your `MS_GRAPH_CLIENT_ID`
7. Copy the **Directory (tenant) ID** - this is your `MS_GRAPH_TENANT_ID`

### Required Permissions

Configure these permissions in your Azure AD app:

1. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
2. Add these permissions:
   - `User.Read` - Read user profile
   - `User.ReadBasic.All` - Read all users' basic profiles
   - `Chat.ReadBasic` - Read basic chat properties
   - `Chat.ReadWrite` - Read and write chats
   - `ChatMessage.Read` - Read chat messages
   - `ChatMessage.Send` - Send chat messages
3. Click **Grant admin consent** (if you have admin rights)

### Authenticate

Run the authentication command once to sign in:

```bash
npm run auth
```

This will open a browser window for you to sign in with your Microsoft account.

## ✨ Features

- 💬 **Teams Chat Management** - Search, create, and manage Teams chats
- 📨 **Message Operations** - Send messages, search conversations, get chat history
- 👥 **User Discovery** - Search users and get profile information
- 🔍 **Powerful Search** - Search across all Teams messages with KQL syntax
- 🎯 **Smart Mentions** - @mention users in messages
- 🔒 **OAuth Authentication** - Secure Azure AD authentication flow
- 🎨 **Rich Formatting** - Markdown support in messages

## 🛠️ Available Tools

The server provides **8 MCP tools** for Microsoft Graph operations:

### User Management
- `get_current_user` - Get your own profile information
- `search_users` - Search for users by name or email
- `get_user` - Get detailed information about a specific user

### Chat Operations
- `search_chats` - Search chats by topic or member name
- `create_chat` - Create new 1:1 or group chats

### Message Operations
- `get_chat_messages` - Retrieve messages from a specific chat
- `send_chat_message` - Send messages with Markdown and mentions
- `search_messages` - Search across all Teams messages using KQL

## 💡 Example Queries

- *"Search for chats with John Smith"*
- *"Show me recent messages from the Engineering chat"*
- *"Send a message to the Dev Team chat saying the deployment is complete"*
- *"Search all Teams messages mentioning the Q4 roadmap"*
- *"Create a group chat with Alice, Bob, and Carol about the new project"*
- *"Find all urgent messages from last week"*
- *"Get my user profile information"*

## 🏗️ Development

### From Source

```bash
# Clone and setup
git clone https://github.com/evrimalacan/mcp-microsoft-graph.git
cd mcp-microsoft-graph
npm install

# Set up credentials
cp .env.example .env.local
# Edit .env.local with your CLIENT_ID and TENANT_ID

# Build
npm run build

# Authenticate
npm run auth

# Development mode
npm run dev

# Run tests
npm test
```

### Adding New Tools

1. Create a new tool file in the appropriate domain folder under `src/tools/`
2. Export it from the domain's `index.ts`
3. Register it in `src/tools/index.ts`

See the existing tools for examples and patterns.

## 🐛 Troubleshooting

### Common Issues

**"Authentication failed"**
- Run `npm run auth` to authenticate again
- Verify your CLIENT_ID and TENANT_ID are correct
- Check that your Azure AD app has the required permissions

**"Access forbidden"**
- Ensure your Azure AD app has the necessary permissions granted
- Check if admin consent is required and has been granted
- Verify you're signed in with the correct account

**"Token expired"**
- Run `npm run auth` to refresh your authentication
- Check that the token file path is accessible

**"Chat not found"**
- Verify the chat ID is correct
- Ensure you have access to the chat
- Check that the chat still exists

## 📚 Documentation

- **[Microsoft Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)** - Interactive API testing
- **[Microsoft Graph Documentation](https://learn.microsoft.com/graph)** - Official API reference
- **[Model Context Protocol](https://modelcontextprotocol.io)** - MCP specification

## 🤝 Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Add tests for your changes
4. Commit your changes (`git commit -m 'Add amazing feature'`)
5. Push to the branch (`git push origin feature/amazing-feature`)
6. Open a Pull Request

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details.

## 🌟 Support

- 🐛 **Bug Reports**: [GitHub Issues](https://github.com/evrimalacan/mcp-microsoft-graph/issues)
- 💡 **Feature Requests**: [GitHub Discussions](https://github.com/evrimalacan/mcp-microsoft-graph/discussions)

---

**Built for seamless Microsoft Teams integration with AI assistants**
