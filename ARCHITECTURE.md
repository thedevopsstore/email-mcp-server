# MS365 Email MCP Server Architecture (Token-Only)

This document describes the **token-only** deployment using `ms365_email_mcp_server/server_token_only.py`.

---

## 1. High-Level Purpose

- Exposes Microsoft 365 Outlook email operations (list folders/messages, send mail, create drafts, delete) via the Model Context Protocol (MCP).
- Designed for autonomous agents: **agents obtain Microsoft Graph access tokens** and pass them to this server.
- Runs over FastMCP using **streamable HTTP** transport (`streamable-http`) so any MCP client can connect via HTTP/SSE.
- Supports shared mailboxes by allowing a `user_identifier` (UserPrincipalName or Graph user ID) to be supplied per request, via custom header, or via env var.
- Avoids handling client secrets inside the server.

---

## 2. Runtime Stack

| Layer        | Technology | Purpose |
|--------------|------------|---------|
| Transport    | FastMCP (`streamable-http`) | MCP runtime and HTTP/SSE transport |
| Auth (Graph) | External (Agent/MSAL) | Agent acquires Graph token; server forwards it to Graph |
| HTTP Client  | `httpx.AsyncClient` | Async Microsoft Graph requests |
| Target API   | Microsoft Graph v1.0 | Mail endpoints |
| Cloud Modes  | Commercial & GCC | Switches Graph base URLs to `.us` when `MS365_CLOUD_TYPE` indicates gov |

---

## 3. Project Layout

```
email-mcp-server/
├── ms365_email_mcp_server/
│   ├── __init__.py
│   ├── server.py              # MSAL client-credentials implementation (legacy/default entrypoint)
│   └── server_token_only.py   # Token-only implementation (Authorization pass-through)
├── ARCHITECTURE.md
├── AUTHENTICATION_FLOW.md
├── MULTI_AGENT_SETUP.md
├── pyproject.toml
└── ...
```

---

## 4. Authentication & Header Flow (AgentCore Runtime)

### 4.1 What the agent sends

- **Graph token**: via standard header
  - `Authorization: Bearer <graph_access_token>`
- **Shared mailbox selector (optional)**: via custom header
  - `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier: <UPN or Graph user id>`

### 4.2 What AgentCore must allowlist

AgentCore only forwards selected headers to the runtime. Ensure your runtime is configured with:

- `Authorization`
- `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier`

The allowlist supports `Authorization` and custom headers prefixed with `X-Amzn-Bedrock-AgentCore-Runtime-Custom-*` (see AWS docs: [RequestHeaderConfiguration](https://docs.aws.amazon.com/bedrock-agentcore-control/latest/APIReference/API_RequestHeaderConfiguration.html)).

### 4.3 What the server does

`server_token_only.py`:
- extracts the token from `Authorization`
- extracts `user_identifier` from the custom header (or tool arg / env var)
- calls Microsoft Graph with `Authorization: Bearer <token>`

---

## 5. `server_token_only.py` Walkthrough

### 5.1 Stateless request handling

`stateless_http=true` means each HTTP request is handled independently; no session affinity is required.

### 5.2 Header extraction

`extract_request_auth_from_headers()` reads both values in one pass:
- token from `authorization`
- shared mailbox from `x-amzn-bedrock-agentcore-runtime-custom-ms365-useridentifier`

### 5.3 Client creation

`get_client()` merges inputs with this priority:
1. tool params (`access_token`, `user_identifier`)
2. request headers
3. env var (`MS365_USER_IDENTIFIER`) for user_identifier

It returns a **new** `MS365EmailClient` each call (no token caching).

### 5.4 Graph API calls

`MS365EmailClient`:
- rewrites `me/...` to `users/{user_identifier}/...` if a user identifier is present
- uses `httpx.AsyncClient` to call Graph
- treats `401` as “token expired/invalid” (agent should refresh)

---

## 6. Environment Variables (Token-Only)

| Variable | Description | Required | Default |
|----------|-------------|----------|---------|
| `MS365_USER_IDENTIFIER` | Default shared mailbox identifier (UPN or Graph ID) | Optional | — |
| `MS365_CLOUD_TYPE` | `commercial`, `gov`, `government`, `usgov` | Optional | `commercial` |
| `HOST`, `PORT`, `STATELESS_HTTP`, `LOG_LEVEL` | FastMCP runtime options | Optional | `0.0.0.0`, `8100`, `true`, `INFO` |

Notes:
- Token-only mode does **not** require `MS365_CLIENT_ID`, `MS365_CLIENT_SECRET`, or `MS365_TENANT_ID`.
- If you later enable JWT verification (FastMCP `JWTVerifier`), you’ll introduce issuer/audience configuration and likely tenant/app ids.



