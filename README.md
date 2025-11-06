# MS365 Email MCP Server

A Python-based Model Context Protocol (MCP) server for Microsoft 365 Outlook email operations. This server uses **OAuth 2.0 Client Credentials Flow** for authentication, making it ideal for autonomous agents and server-to-server scenarios.

Built with [FastMCP](https://github.com/jlowin/fastmcp) for simplified server implementation, following the pattern used by [AWS API MCP Server](https://github.com/awslabs/mcp).

## Features

This MCP server provides the following email operations:

- **list-mail-messages** - List mail messages from inbox or a specific folder
- **list-mail-folders** - List all mail folders
- **list-mail-folder-messages** - List messages from a specific folder
- **get-mail-message** - Get a specific mail message by ID
- **send-mail** - Send an email
- **delete-mail-message** - Delete a mail message
- **create-draft-email** - Create a draft email
- **move-mail-message** - Move a mail message to another folder

## Authentication

This server uses **OAuth 2.0 Client Credentials Flow** (app-only authentication), which is perfect for autonomous agents and background services. No user interaction is required.

**Reference**: [Microsoft Graph - Get access without a user](https://learn.microsoft.com/en-us/graph/auth-v2-service?tabs=http)

## Prerequisites

### 1. Azure App Registration

You need to register an application in Azure AD with **Application permissions** (not Delegated):

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Create a new registration
3. Note down:
   - **Application (client) ID**
   - **Directory (tenant) ID**
4. Create a **client secret**:
   - Go to Certificates & secrets → New client secret
   - Copy the secret value immediately (it won't be shown again)
5. Configure **API permissions**:
   - Go to API permissions → Add a permission → Microsoft Graph
   - Select **Application permissions** (not Delegated)
   - Add the following permissions:
     - `Mail.Read` - Read user mail
     - `Mail.ReadWrite` - Read and write user mail
     - `Mail.Send` - Send mail as user
   - **Grant admin consent** (required for application permissions)

**Important**: Application permissions require administrator consent. The app must be granted consent before it can use these permissions.

**Reference**: [Microsoft Graph - Configure permissions](https://learn.microsoft.com/en-us/graph/auth-v2-service?tabs=http#step-1-configure-permissions-for-microsoft-graph)

### 2. Environment Variables

```bash
MS365_CLIENT_ID=your-azure-ad-client-id
MS365_CLIENT_SECRET=your-azure-ad-client-secret
MS365_TENANT_ID=your-tenant-id
# Optional: For Government Cloud
MS365_CLOUD_TYPE=gov  # Options: "commercial" (default), "gov", "government", "usgov"
```

## Local Development

This project uses [UV](https://github.com/astral-sh/uv) for dependency management, similar to the [AWS API MCP Server](https://github.com/awslabs/mcp/tree/main/src/aws-api-mcp-server).

### 1. Install UV

```bash
# Install UV using pip
pip install uv

# Or using Homebrew (macOS)
brew install uv

# Or using the official installer
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### 2. Install Dependencies

```bash
# Install project dependencies
uv sync

# Or install in development mode with dev dependencies
uv sync --dev
```

### 3. Set Environment Variables

```bash
export MS365_CLIENT_ID="your-client-id"
export MS365_CLIENT_SECRET="your-client-secret"
export MS365_TENANT_ID="your-tenant-id"
export PORT="8100"  # Optional, defaults to 8100
export HOST="0.0.0.0"  # Optional, defaults to 0.0.0.0
export LOG_LEVEL="INFO"  # Optional, defaults to INFO
export STATELESS_HTTP="true"  # Optional, defaults to true
```

### 4. Run the Server

Using UV (recommended):

```bash
# Run using UV
uv run ms365-email-mcp-server

# Or run directly with Python (after uv sync)
python -m ms365_email_mcp_server.server
```

The server will start on `http://localhost:8100` with:
- **SSE endpoint**: `http://localhost:8100/message`
- **Health check**: `http://localhost:8100/health`

## Docker Usage

### Build the Docker Image

The Dockerfile uses UV for dependency management:

```bash
docker build -t email-mcp-server .
```

### Run the Container

```bash
docker run -d \
  -p 8100:8100 \
  -e MS365_CLIENT_ID="your-client-id" \
  -e MS365_CLIENT_SECRET="your-client-secret" \
  -e MS365_TENANT_ID="your-tenant-id" \
  -e MS365_CLOUD_TYPE="commercial" \
  email-mcp-server
```

### Using Docker Compose

Create a `.env` file:

```bash
MS365_CLIENT_ID=your-client-id
MS365_CLIENT_SECRET=your-client-secret
MS365_TENANT_ID=your-tenant-id
MS365_CLOUD_TYPE=commercial
```

Then run:

```bash
docker-compose up
```

Or in detached mode:

```bash
docker-compose up -d
```

## MCP Client Configuration

### HTTP/SSE Endpoint

The server runs as an HTTP/SSE server. Connect to it using:

- **SSE Endpoint**: `http://localhost:8100/message`
- **Health Check**: `http://localhost:8100/health`

### Using with MCP Clients

#### Option 1: HTTP/SSE Transport (Recommended)

For clients that support HTTP/SSE transport, configure:

```json
{
  "mcpServers": {
    "ms365-email": {
      "url": "http://localhost:8100/message",
      "transport": "sse"
    }
  }
}
```

#### Option 2: Using UV (for stdio transport)

If you want to use UV to run the server (similar to AWS API MCP Server), you can configure:

```json
{
  "mcpServers": {
    "ms365-email": {
      "command": "uvx",
      "args": [
        "ms365-email-mcp-server@latest"
      ],
      "env": {
        "MS365_CLIENT_ID": "your-client-id",
        "MS365_CLIENT_SECRET": "your-client-secret",
        "MS365_TENANT_ID": "your-tenant-id",
        "LOG_LEVEL": "INFO"
      }
    }
  }
}
```

**Note**: This requires the package to be published to PyPI. For local development, use Option 1 or run the server manually with `uv run ms365-email-mcp-server`.

## API Reference

### Microsoft Graph API Endpoints Used

- **List Messages**: [GET /me/messages](https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http)
- **Send Mail**: [POST /me/sendMail](https://learn.microsoft.com/en-us/graph/api/user-post-messages?view=graph-rest-1.0&tabs=http)
- **Get Message**: `GET /me/messages/{id}`
- **Delete Message**: `DELETE /me/messages/{id}`
- **Create Draft**: `POST /me/messages`
- **Move Message**: `POST /me/messages/{id}/move`
- **List Folders**: `GET /me/mailFolders`

## Government Cloud Support

For Azure Government Cloud, set:

```bash
export MS365_CLOUD_TYPE="gov"
```

This will use:
- Authority: `https://login.microsoftonline.us`
- Graph API: `https://graph.microsoft.us`

## Security Notes

- Never commit your client secrets to version control
- Use environment variables or secret management systems
- Rotate client secrets regularly
- Consider using managed identities in Azure for production
- Application permissions require administrator consent

## Troubleshooting

### "Failed to acquire token"

- Verify your `MS365_CLIENT_ID`, `MS365_CLIENT_SECRET`, and `MS365_TENANT_ID` are correct
- Ensure admin consent has been granted for the application permissions
- Check that the client secret hasn't expired

### "Insufficient privileges"

- Verify that admin consent has been granted for all required permissions
- Check that you're using **Application permissions** (not Delegated permissions)

### "Invalid tenant"

- Verify your `MS365_TENANT_ID` is correct
- For Government Cloud, ensure `MS365_CLOUD_TYPE` is set to `"gov"`

## License

MIT

