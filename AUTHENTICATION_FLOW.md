# Authentication Flow: MCP Server on AWS AgentCore Runtime

This document explains the dual authentication layers when hosting an MCP server on AWS AgentCore Runtime:

1. **Inbound Authentication**: How clients authenticate to the MCP server (OAuth 2.0)
2. **Outbound Authentication**: How the MCP server authenticates to third-party providers (Microsoft Graph API, etc.)

---

## Overview: Two-Layer Authentication

```
┌─────────────┐
│   Client    │ (Agent or User)
│  (OAuth 2.0 │
│   JWT Token)│
└──────┬──────┘
       │ 1. Inbound Auth
       │    (OAuth 2.0 JWT)
       ▼
┌─────────────────────────────────────┐
│   AWS AgentCore Runtime              │
│   - Validates JWT token              │
│   - Routes to MCP server             │
│   - Provides Workload Access Token   │
└──────┬──────────────────────────────┘
       │
       │ 2. MCP Server receives request
       ▼
┌─────────────────────────────────────┐
│   MCP Server                        │
│   (Your MS365 Email MCP Server)     │
└──────┬──────────────────────────────┘
       │
       │ 3. Outbound Auth
       │    (Authenticate to Microsoft)
       ▼
┌─────────────────────────────────────┐
│   Third-Party Provider              │
│   (Microsoft Graph API)             │
└─────────────────────────────────────┘
```

---

## Layer 1: Inbound Authentication (Client → MCP Server)

### How It Works

When a client (agent or user) wants to access your MCP server hosted on AgentCore Runtime:

1. **Client obtains OAuth 2.0 JWT token** from an identity provider (e.g., AWS Cognito, Auth0, Microsoft Entra ID)
2. **Client sends request** to AgentCore Runtime with JWT token in `Authorization: Bearer <token>` header
3. **AgentCore Runtime validates the JWT token** using the configured authorizer:
   - Validates token signature
   - Checks `iss` (issuer) matches discovery URL
   - Verifies `aud` (audience) is in allowed audiences
   - Validates `client_id` is in allowed clients
   - Checks `scope` matches allowed scopes
   - Verifies token expiration
4. **AgentCore Runtime routes request** to your MCP server at `0.0.0.0:8000/mcp`
5. **AgentCore Runtime provides Workload Access Token** in the `WorkloadAccessToken` header for your MCP server to use

### Configuration

When creating your AgentCore Runtime, configure the inbound authorizer:

```bash
aws bedrock-agentcore-control create-agent-runtime \
  --agent-runtime-name ms365-email-mcp-server \
  --inbound-authorization-configuration '{
    "customJwtAuthorizerConfiguration": {
      "discoveryUrl": "https://login.microsoftonline.com/{tenant_id}/.well-known/openid-configuration",
      "allowedClients": ["client-id-1", "client-id-2"],
      "allowedAudiences": ["api://your-api"],
      "allowedScopes": ["read:emails", "send:emails"]
    }
  }'
```

**Key Points:**
- AgentCore Runtime handles all JWT validation
- Your MCP server doesn't need to validate inbound tokens
- Multiple clients can authenticate (multi-agent support)
- OAuth 2.0 compliant (supports standard JWT tokens)

---

## Layer 2: Outbound Authentication (MCP Server → Third-Party Providers)

This is where your MCP server authenticates to Microsoft Graph API (or other third-party providers). There are **two main approaches**:

### Approach 1: AgentCore Identity Service (Recommended for Production)

**How It Works:**

1. **Register OAuth Credential Provider** with AgentCore Identity:
   ```bash
   aws bedrock-agentcore-control create-oauth2-credential-provider \
     --name "microsoft-graph-provider" \
     --credential-provider-vendor "Microsoft" \
     --oauth2-provider-config-input '{
       "microsoftProviderConfig": {
         "clientId": "your-azure-app-client-id",
         "clientSecret": "your-azure-app-client-secret",
         "tenantId": "your-tenant-id"
       }
     }'
   ```

2. **MCP Server uses AgentCore SDK** to request tokens:
   ```python
   from bedrock_agentcore.identity.auth import requires_access_token
   
   @requires_access_token(
       provider_name="microsoft-graph-provider",
       scopes=["https://graph.microsoft.com/.default"],
       auth_flow="CLIENT_CREDENTIALS",  # For app-only auth
       force_authentication=True
   )
   async def call_graph_api(*, access_token: str):
       # Use access_token to call Microsoft Graph API
       headers = {"Authorization": f"Bearer {access_token}"}
       # ... make API call
   ```

3. **AgentCore Identity Service**:
   - Manages OAuth 2.0 flows (client credentials, authorization code, etc.)
   - Securely stores tokens in Token Vault
   - Automatically refreshes expired tokens
   - Binds tokens to workload identity + user ID

**Benefits:**
- ✅ Centralized credential management
- ✅ Automatic token refresh
- ✅ Secure token storage (encrypted at rest and in transit)
- ✅ No secrets in your code
- ✅ Supports both client credentials and authorization code flows

**What Happens Behind the Scenes:**
1. MCP server receives `WorkloadAccessToken` from AgentCore Runtime
2. MCP server calls `bedrock-agentcore:GetResourceOauth2Token` API with Workload Access Token
3. AgentCore Identity checks Token Vault for existing valid token
4. If no valid token exists, AgentCore Identity performs OAuth flow:
   - For **client credentials flow**: Exchanges client_id/secret for token
   - For **authorization code flow**: Generates auth URL, user grants consent, exchanges code for token
5. Token is stored in Token Vault and returned to MCP server
6. MCP server uses token to call Microsoft Graph API

---

### Approach 2: Token Pass-Through (Current Implementation)

**How It Works:**

This is what your current `server_token_only.py` implementation does:

1. **Agents obtain Microsoft Graph tokens themselves** (using their own Azure AD app registrations)
2. **Agents pass the Graph token via the standard `Authorization` header** (passed through to the runtime):
   ```python
   headers = {
       "Authorization": f"Bearer {graph_token}"
   }
   ```

3. **MCP server extracts token from headers** and uses it directly:
   ```python
   def extract_token_from_headers() -> Optional[str]:
       headers = get_http_headers(include_all=True)
       auth = headers.get("authorization", "")
       return auth[7:] if auth.lower().startswith("bearer ") else None
   ```

4. **MCP server uses token directly** for Microsoft Graph API calls (no token exchange)

**Benefits:**
- ✅ Simple implementation
- ✅ No AgentCore Identity setup required
- ✅ Each agent uses its own credentials (multi-tenant)
- ✅ Agents control token refresh

**Limitations:**
- ⚠️ Agents must manage token refresh themselves
- ⚠️ Tokens appear in headers (though encrypted in transit)
- ⚠️ No centralized token management

---

## Comparison: AgentCore Identity vs Token Pass-Through

| Aspect | AgentCore Identity | Token Pass-Through |
|--------|-------------------|-------------------|
| **Setup Complexity** | Higher (requires credential provider setup) | Lower (just pass tokens) |
| **Token Management** | Automatic (AgentCore handles refresh) | Manual (agents handle refresh) |
| **Security** | Tokens stored in encrypted vault | Tokens in headers (encrypted in transit) |
| **Multi-Agent Support** | Yes (via user ID binding) | Yes (each agent passes own token) |
| **Token Refresh** | Automatic | Manual (agents must refresh) |
| **Best For** | Production, centralized management | Development, multi-tenant scenarios |

---

## Current Implementation Analysis

Your `server_token_only.py` uses **Approach 2 (Token Pass-Through)**:

1. **Inbound**: AgentCore Runtime validates JWT tokens from agents
2. **Outbound**: MCP server extracts Microsoft Graph tokens from custom headers and uses them directly

**Flow:**
```
Agent → [JWT Token] → AgentCore Runtime → [Validates JWT] → MCP Server
                                                              ↓
Agent → [Graph Token in Header] → AgentCore Runtime → [Passes Header] → MCP Server
                                                              ↓
MCP Server → [Extracts Graph Token] → [Uses Token] → Microsoft Graph API
```

---

## Migration Path: Token Pass-Through → AgentCore Identity

If you want to migrate to Approach 1 (AgentCore Identity):

1. **Register Microsoft OAuth Credential Provider**:
   ```bash
   aws bedrock-agentcore-control create-oauth2-credential-provider \
     --name "microsoft-graph-provider" \
     --credential-provider-vendor "Microsoft" \
     --oauth2-provider-config-input '{...}'
   ```

2. **Update MCP Server Code**:
   - Add `bedrock-agentcore` SDK dependency
   - Replace token extraction with `@requires_access_token` decorator
   - Remove manual token handling

3. **Update Agent Code**:
   - Remove token acquisition logic (AgentCore handles it)
   - Remove custom header passing

---

## Security Considerations

### Inbound Authentication (Client → MCP Server)
- ✅ JWT tokens validated by AgentCore Runtime (signature, expiration, claims)
- ✅ Only allowed clients can authenticate
- ✅ Tokens encrypted in transit (HTTPS)

### Outbound Authentication (MCP Server → Third-Party)

**With AgentCore Identity:**
- ✅ Tokens stored in encrypted Token Vault
- ✅ Automatic token refresh
- ✅ Access controlled by IAM policies
- ✅ Audit logging

**With Token Pass-Through:**
- ✅ Tokens encrypted in transit (HTTPS)
- ⚠️ Agents must securely store client secrets
- ⚠️ Agents must handle token refresh
- ⚠️ Tokens visible in headers (though encrypted)

---

## Summary

**Question**: "How does the MCP server authenticate to third-party providers?"

**Answer**: There are two approaches:

1. **AgentCore Identity Service** (Recommended):
   - MCP server uses AgentCore SDK to request tokens
   - AgentCore Identity manages OAuth flows and token storage
   - Tokens are automatically refreshed
   - No secrets in your code

2. **Token Pass-Through** (Current Implementation):
   - Agents obtain tokens themselves
   - Agents pass tokens via custom headers
   - MCP server uses tokens directly
   - Agents manage token refresh

Both approaches work, but AgentCore Identity provides better security, centralized management, and automatic token refresh for production deployments.

