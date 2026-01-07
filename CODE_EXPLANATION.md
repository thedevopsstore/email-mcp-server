# Complete Code Explanation: MS365 Email MCP Server (Token-Only)

This document explains the **token-only** implementation in `ms365_email_mcp_server/server_token_only.py`.

In this model, the **agent obtains a Microsoft Graph access token** (for example using MSAL) and passes it to the MCP server via an AgentCore Runtime custom header:

- `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization: Bearer <graph_access_token>`

The server uses that token directly to call Microsoft Graph; it does **not** accept client secrets and does **not** perform token acquisition/refresh.

---

## Table of Contents

1. [Overall Architecture](#overall-architecture)
2. [Configuration & Transport](#configuration--transport)
3. [Header Extraction](#header-extraction)
4. [`MS365EmailClient`](#ms365emailclient)
5. [`get_client`](#get_client)
6. [MCP Tools](#mcp-tools)
7. [Environment Variables](#environment-variables)
8. [Deployment Notes (AgentCore Header Pass-Through)](#deployment-notes-agentcore-header-pass-through)
9. [Resources & Prompts](#resources--prompts)

---

## Overall Architecture

```
┌─────────────────────────────────────────┐
│ FastMCP Server (MCP Protocol)           │
├─────────────────────────────────────────┤
│ Tool Functions (MCP wrappers)           │
├─────────────────────────────────────────┤
│ MS365EmailClient (Graph API logic)      │
├─────────────────────────────────────────┤
│ httpx.AsyncClient (HTTP)                │
└─────────────────────────────────────────┘
                    │
                    ▼
          Microsoft Graph (v1.0)
```

Key properties:
- **Token-only**: no MSAL usage in the server
- **Stateless**: each request can be handled independently (`stateless_http=True`)
- **Shared mailbox aware**: optional `user_identifier` rewrites `/me/...` to `/users/{id}/...`

---

## Configuration & Transport

`server_token_only.py` configures:
- `TRANSPORT = "streamable-http"`
- `STATELESS_HTTP = true` by default

This makes the server suitable for horizontal scaling without sticky sessions.

---

## Header Extraction

`extract_request_auth_from_headers()` reads both token and shared-mailbox selector from HTTP headers (via FastMCP’s `get_http_headers()`):

- **Graph token** from `x-amzn-bedrock-agentcore-runtime-custom-ms365-authorization` (expects `Bearer ...` but also accepts raw token)
- **Shared mailbox selector** from:
  - `x-amzn-bedrock-agentcore-runtime-custom-ms365-useridentifier`
  - which corresponds to the original header name `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier`

It returns:
- `(access_token: Optional[str], user_identifier: Optional[str])`

---

## `MS365EmailClient`

`MS365EmailClient(access_token, user_identifier=None, cloud_type=None)`:

- Stores the token (strips `"Bearer "` prefix if present)
- Sets Graph base URL:
  - `https://graph.microsoft.com` (commercial)
  - `https://graph.microsoft.us` (gov, when `MS365_CLOUD_TYPE` is set accordingly)
- Applies shared-mailbox rewrite:
  - `_build_endpoint("me/…")` → `"users/{user_identifier}/…"` when `user_identifier` exists

`_make_request()`:
- ensures endpoint has `/v1.0/` prefix
- calls Graph with `Authorization: Bearer <token>`
- warns on `401` (caller should refresh token)

---

## `get_client`

`get_client(access_token=None, user_identifier=None, cloud_type=None)` merges values in this priority:

- **access_token**:
  - tool parameter → `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization` header
- **user_identifier**:
  - tool parameter → custom header → `MS365_USER_IDENTIFIER` env var

It returns a **new `MS365EmailClient` per call**.

Rationale:
- avoids storing bearer tokens in global caches
- keeps behavior deterministic and request-scoped in stateless mode

---

## MCP Tools

Tools are thin wrappers around `MS365EmailClient` methods:

- `list-mail-messages` (preview only)
- `list-mail-folders`
- `get-mail-message` (full message; marks read)
- `send-mail`
- `delete-mail-message`
- `create-draft-email`

Common pattern:
- `access_token` is optional in tool args because it can come from `Authorization` header.
- `user_identifier` is optional in tool args because it can come from the custom header.

---

## Environment Variables

Token-only mode does **not** require client credentials.

| Variable | Description | Required | Default |
|----------|-------------|----------|---------|
| `MS365_CLOUD_TYPE` | `commercial`, `gov`, `government`, `usgov` | Optional | `commercial` |
| `MS365_USER_IDENTIFIER` | Default shared mailbox identifier (UPN or Graph ID) | Optional | — |
| `HOST`, `PORT`, `STATELESS_HTTP`, `LOG_LEVEL` | FastMCP runtime options | Optional | `0.0.0.0`, `8100`, `true`, `INFO` |

---

## Deployment Notes (AgentCore Header Pass-Through)

To use `Authorization` inside the MCP server, your AgentCore runtime must allowlist it for pass-through.

- AWS docs: [RequestHeaderConfiguration](https://docs.aws.amazon.com/bedrock-agentcore-control/latest/APIReference/API_RequestHeaderConfiguration.html)

Typical allowlist:
- `Authorization`
- `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier` (if you want shared mailbox selection via header)

---

## Resources & Prompts

This server is action-oriented (send/list/get/delete email), so **Tools** are the right MCP primitive.

- Resources (data exposure) and Prompts (templated LLM messages) aren’t required for this server’s core functionality.


