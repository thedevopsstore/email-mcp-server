# Multi-Agent Setup for MS365 Email MCP Server

This document explains how to configure the MS365 Email MCP Server to support multiple agents, where each agent uses its own Microsoft Entra ID credentials to access Microsoft Graph API.

## Architecture Overview

```
┌─────────────┐
│ Agent 1     │ ──┐
│ (Client ID 1)│   │
└─────────────┘   │
                  │
┌─────────────┐   │    ┌─────────────────────────┐
│ Agent 2     │ ──┼───▶│ AgentCore Runtime       │
│ (Client ID 2)│   │    │ - Validates Agent JWTs  │
└─────────────┘   │    │ - Routes to MCP Server  │
                  │    │ - Passes custom headers │
┌─────────────┐   │    └──────────┬──────────────┘
│ Agent 3     │ ──┘               │
│ (Client ID 3)│                  │
└─────────────┘                   │
                                  ▼
                    ┌─────────────────────────┐
                    │ MCP Server              │
                    │ - Extracts credentials   │
                    │   from headers/params    │
                    │ - Uses agent's creds to  │
                    │   call Microsoft Graph  │
                    └──────────┬──────────────┘
                               │
                               ▼
                    ┌─────────────────────────┐
                    │ Microsoft Graph API     │
                    │ (Each agent uses its own │
                    │  app registration)       │
                    └─────────────────────────┘
```

## Solution: Pass Microsoft Graph Token via AgentCore Custom Header (Pass-Through)

AgentCore Runtime can be configured to pass through an AgentCore custom header to your MCP server. Each agent obtains its own Microsoft Graph API token and passes it in:

- `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization: Bearer <token>`

This avoids using the inbound `Authorization` header for Graph tokens and works well when inbound authentication is IAM-based.

### Step 1: Configure AgentCore Runtime Header Allowlist

When creating or updating your AgentCore Runtime, configure it to allow the custom headers you need (Graph token + shared-mailbox `user_identifier`):

```bash
aws bedrock-agentcore-control create-agent-runtime \
  --agent-runtime-name ms365-email-mcp-server \
  --agent-runtime-artifact containerConfiguration={containerUri=your-ecr-uri} \
  --request-header-configuration '{
    "requestHeaderAllowlist": [
      "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization",
      "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier"
    ]
  }'
```

**Important**: Custom headers must start with `X-Amzn-Bedrock-AgentCore-Runtime-Custom-`

### Inbound Authentication Note

If you are using **IAM-based access** to AgentCore Runtime (recommended for your scenario), you do **not** configure a JWT authorizer. Each agent/service calls the runtime using AWS credentials (SigV4).

### Step 2: Agents Obtain and Pass Token

Each agent:
1. Uses its own `client_id` and `client_secret` to get a JWT token for Microsoft Graph API
2. Passes that token in the `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization` header

**Helper function to get Microsoft Graph API token:**

```python
import requests
import os
from typing import Optional

def get_graph_api_token(
    client_id: str,
    client_secret: str,
    tenant_id: str,
    cloud_type: str = "commercial"
) -> str:
    """
    Get Microsoft Graph API access token using client credentials flow.
    
    Args:
        client_id: Microsoft Entra app client ID
        client_secret: Microsoft Entra app client secret
        tenant_id: Microsoft Entra tenant ID
        cloud_type: "commercial" or "gov" (default: "commercial")
    
    Returns:
        Access token (JWT) for Microsoft Graph API
    """
    authority_base = (
        "https://login.microsoftonline.us" 
        if cloud_type.lower() in ["gov", "government", "usgov"] 
        else "https://login.microsoftonline.com"
    )
    
    token_url = f"{authority_base}/{tenant_id}/oauth2/v2.0/token"
    
    response = requests.post(
        token_url,
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials"
        }
    )
    response.raise_for_status()
    return response.json()["access_token"]
```

**Using the token in agent code:**

```python
# In your Strands agent code
import requests

# Step 2: Get Microsoft Graph API token using agent's credentials
graph_token = get_graph_api_token(
    client_id=os.getenv("AGENT_CLIENT_ID"),
    client_secret=os.getenv("AGENT_CLIENT_SECRET"),
    tenant_id=os.getenv("AGENT_TENANT_ID")
)

# Step 3: Invoke MCP server with Graph token in custom header
response = requests.post(
    f"https://bedrock-agentcore.{region}.amazonaws.com/runtimes/{runtime_arn}/invocations",
    headers={
        # Inbound auth to AgentCore Runtime is IAM (SigV4) in this model.
        # Use AWS SDK / SigV4-capable HTTP client in production.
        "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization": f"Bearer {graph_token}",
        "Content-Type": "application/json"
    },
    json={
        "prompt": "List my emails"
    }
)
```

### Step 3: MCP Server Uses Token Directly

The MCP server:
1. Extracts the token from the custom header
2. Uses it directly for Microsoft Graph API calls (no token exchange needed)
3. Each agent's token is cached separately

**Benefits of this approach:**
- ✅ No client secrets in headers (more secure)
- ✅ Tokens are already obtained by agents
- ✅ MCP server just uses tokens directly
- ✅ Supports multiple agents with different tokens

## Alternative: Pass Token via Tool Parameters

If header extraction is not available in your FastMCP version, agents can pass the token as a tool parameter:

```python
# Agent gets Graph API token first
graph_token = get_graph_api_token(
    client_id="agent-1-client-id",
    client_secret="agent-1-secret",
    tenant_id="tenant-id"
)

# Agent calls MCP tool with token
result = await mcp_client.call_tool(
    "list-mail-messages",
    {
        "top": 10,
        "access_token": graph_token  # Bearer token for Microsoft Graph API
    }
)
```

**Note**: This approach is less secure as tokens appear in tool call logs. Prefer headers when possible.

## Security Considerations

1. **No secrets in headers**: Only tokens are passed, never client secrets
2. **Token expiration**: Tokens expire (typically 60-90 minutes), agents must refresh them
3. **Use HTTPS**: All communication must be over HTTPS
4. **Least privilege**: Each agent's app registration should have only necessary permissions
5. **Header validation**: AgentCore Runtime validates agent JWTs before passing headers to MCP server
6. **Token scope**: Ensure tokens are scoped to `https://graph.microsoft.com/.default` or specific Graph API scopes

## Client Caching

The MCP server caches `MS365EmailClient` instances per `(token_hash, user_identifier)` tuple. This means:
- Each agent (with different tokens) gets its own cached client
- Token-based clients don't need MSAL (token is used directly)
- Multiple agents can use the server simultaneously
- Tokens are not cached by the server (agents manage token refresh)

## Testing

To test with multiple agents:

1. Create multiple Microsoft Entra app registrations (one per agent)
2. Configure each with appropriate Microsoft Graph permissions
3. Deploy MCP server to AgentCore Runtime with header allowlist
4. Each agent passes its own credentials via headers
5. Verify each agent accesses its own mailbox/permissions

## Troubleshooting

### Headers not reaching MCP server
- Verify header allowlist in AgentCore Runtime configuration
- Check header names match exactly (case-sensitive)
- Ensure headers start with `X-Amzn-Bedrock-AgentCore-Runtime-Custom-`

### Token not working
- Verify each agent's app registration has correct permissions
- Check tokens are scoped to `https://graph.microsoft.com/.default`
- Ensure tokens are not expired (typically valid for 60-90 minutes)
- Verify token format: should be a JWT, with or without "Bearer " prefix

### Token acquisition fails
- Check Microsoft Graph permissions are granted and admin-consented
- Verify app registrations are in the same tenant (or multi-tenant configured)
- Check cloud type matches (commercial vs government)

