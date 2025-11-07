# MS365 Email MCP Server Architecture

This document explains how the `ms365_email_mcp_server` works end-to-end: runtime stack, authentication flow, shared-mailbox support, transport configuration, and the individual code components.

---

## 1. High-Level Purpose

- Exposes Microsoft 365 Outlook email operations (list folders/messages, send mail, create drafts, delete, move) via the Model Context Protocol (MCP).
- Designed for autonomous agents: authenticates with Azure AD (Microsoft Entra ID) using the client credentials flow (service principal).
- Runs over FastMCP using **streamable HTTP** transport (`streamable-http`) so any MCP client can connect via HTTP/SSE.
- Supports shared mailboxes by allowing a `user_identifier` (UserPrincipalName or Graph user ID) to be supplied per request or via environment variable.
- Handles Microsoft Graph token management automatically via MSAL (token caching, refresh, retry on 401).

---

## 2. Runtime Stack

| Layer        | Technology | Purpose |
|--------------|------------|---------|
| Transport    | FastMCP (`streamable-http`) | MCP runtime and HTTP/SSE transport |
| Auth         | MSAL Python | Acquires Azure AD tokens via confidential client app ([MSAL repo](https://github.com/AzureAD/microsoft-authentication-library-for-python)) |
| HTTP Client  | `httpx.AsyncClient` | Async Microsoft Graph requests |
| Target API   | Microsoft Graph v1.0 | Mail endpoints ([SendMail](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http), [List Messages](https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http)) |
| Cloud Modes  | Commercial & GCC | Switches Graph base URLs to `.us` when `MS365_CLOUD_TYPE` indicates government cloud |

---

## 3. Project Layout

```
email-mcp-server/
├── ms365_email_mcp_server/
│   ├── __init__.py
│   └── server.py          # Main implementation
├── ARCHITECTURE.md        # (this file)
├── README.md              # Setup & usage instructions
├── pyproject.toml         # UV project configuration (entrypoint script)
├── Dockerfile             # UV-based container build
└── Makefile               # Developer convenience commands
```

---

## 4. `server.py` Walkthrough

### 4.1 Configuration & Transport

```python
HOST = os.getenv("HOST", "0.0.0.0")
PORT = int(os.getenv("PORT", "8100"))
STATELESS_HTTP = os.getenv("STATELESS_HTTP", "true").lower() == "true"
TRANSPORT = "streamable-http"
server = FastMCP(...)
```

- Hardcodes `streamable-http` to mirror the AWS API MCP server pattern for HTTP/SSE.
- `stateless_http` (default true) keeps each request isolated.

### 4.2 `MS365EmailClient`

Handles authentication, endpoint selection, and HTTP requests.

#### Initialization

- Reads `MS365_CLIENT_ID`, `MS365_CLIENT_SECRET`, `MS365_TENANT_ID` (required).
- Detects cloud type (commercial vs GCC) via `MS365_CLOUD_TYPE` to choose the correct Graph base.
- Accepts an optional `user_identifier` (UPN or Graph user ID) either as a constructor argument or from `MS365_USER_IDENTIFIER`. When provided, all `/me/...` endpoints are rewritten to `/users/{identifier}/...` so the client can act on shared mailboxes.
- Creates a single `ConfidentialClientApplication` (MSAL) reused for the lifetime of the process.

#### Token Acquisition

```python
result = self.app.acquire_token_silent(self.scope, account=None)
if not result:
    result = self.app.acquire_token_for_client(scopes=self.scope)
```

- Follows MSAL best practice: try the cache first, then fetch.
- MSAL stores client-credential tokens in memory; no manual expiry tracking is required.

#### Graph Request Helper (`_make_request`)

- Rewrites `/me/` to `/users/{identifier}/` when `user_identifier` is set.
- Guarantees endpoints start with `/v1.0` or `/beta` before constructing the full URL.
- Performs each HTTP request inside a 2-attempt loop:
  1. Acquire token (cached or new).
  2. Call Microsoft Graph via `httpx.AsyncClient`.
  3. If Graph returns 401 on the first attempt, clear the MSAL cache and retry once (forces token refresh).
- Handles endpoints returning `202 Accepted` or `204 No Content` by skipping JSON parsing and returning `{"status": code, "status_text": reason}`.
- Logs and raises exceptions for any non-successful responses.

### 4.3 Email Operations

Each method uses `_make_request` and maps to a Graph endpoint:

| Method | Graph Endpoint | Notes |
|--------|----------------|-------|
| `list_mail_messages(folder_id, top)` | `GET me/messages` or `me/mailFolders/{id}/messages` | `top` defaults to 25 |
| `list_mail_folders()` | `GET me/mailFolders` | |
| `list_mail_folder_messages(folder_id, top)` | Calls `list_mail_messages` with folder context | |
| `get_mail_message(message_id)` | `GET me/messages/{id}` | |
| `send_mail(to, subject, body, body_type)` | `POST me/sendMail` | Returns 202 Accepted (queued delivery) |
| `create_draft_email(to, subject, body, body_type)` | `POST me/messages` | |
| `delete_mail_message(message_id)` | `DELETE me/messages/{id}` | Returns 204 No Content |
| `move_mail_message(message_id, destination_id)` | `POST me/messages/{id}/move` | |

> When `user_identifier` is specified, all `/me/...` calls transparently become `/users/{identifier}/...`, allowing the server to operate on shared mailboxes with app-only credentials.

### 4.4 MCP Tool Definitions

Decorated with `@server.tool`, each exposes the mail operations to MCP clients. Example (`send-mail`):

```python
@server.tool(
    name="send-mail",
    description="Send an email...",
    annotations=ToolAnnotations(
        title="Send email",
        readOnlyHint=False,
        destructiveHint=False,
    ),
)
async def send_mail(..., user_identifier: Optional[str] = None, ctx: Context = None) -> dict[str, Any]:
    client = get_client(user_identifier=user_identifier)
    result = await client.send_mail(...)
    return {"success": True, "result": result}
```

- `user_identifier` can be supplied per call (overrides the env var).
- Errors are logged and reported back to the MCP client via `ctx.error`.

### 4.5 Server Entry Point

```python
def main():
    assert MS365_CLIENT_ID/SECRET/TENANT_ID are set
    log settings
    server.run(transport=TRANSPORT)
```

- `transport` is hardcoded to `streamable-http`, matching the AWS MCP reference implementation.
- Called from the `pyproject.toml` entry point (`ms365-email-mcp-server`).

---

## 5. Environment Variables

| Variable | Description | Required | Default |
|----------|-------------|----------|---------|
| `MS365_CLIENT_ID` | Azure AD application (client) ID | ✅ | — |
| `MS365_CLIENT_SECRET` | Client secret (or certificate) | ✅ | — |
| `MS365_TENANT_ID` | Directory/tenant ID | ✅ | — |
| `MS365_USER_IDENTIFIER` | UserPrincipalName or Graph ID for shared mailbox | Optional (required for shared mailboxes) | — |
| `MS365_CLOUD_TYPE` | `commercial`, `gov`, `government`, `usgov` | Optional | `commercial` |
| `MS365_CLOUD_TYPE=gov` | Graph base switches to `https://graph.microsoft.us` | — | — |
| `HOST`, `PORT`, `STATELESS_HTTP`, `LOG_LEVEL` | Runtime/options for FastMCP | Optional | `0.0.0.0`, `8100`, `true`, `INFO` |


---

## 6. Token Handling Summary

1. The first Graph call invokes `acquire_token_silent()` → empty cache → `acquire_token_for_client()` → token cached by MSAL.
2. Subsequent calls reuse the cached token until it expires.
3. If Graph returns `401 Unauthorized`: the MSAL cache is cleared and the request is retried once, fetching a new token from Azure AD.

This flow follows MSAL’s recommended usage pattern, using the library’s built-in cache rather than tracking expiry manually.

---

## 7. Deployment Notes

- **UV Packaging**: Declared in `pyproject.toml` with entry point `ms365-email-mcp-server` and dependencies (`fastmcp`, `msal`, `httpx`, `pydantic`, `loguru`).
- **Docker Image**: Builds with UV (`uv pip install --system -e .`), copies `README.md` earlier because `pyproject.toml` references it.
- **README.md**: Provides environment setup, UV instructions, Docker usage, MCP client configuration (HTTP/SSE and stdio), and shared mailbox guidance.

---

## 8. Key Features Recap

- **Autonomous-agent friendly** (pure app-only auth; no human login).
- **Shared mailbox support** through `user_identifier` and `/users/{id}/` endpoint rewrites.
- **MSAL-backed token caching** with automatic refresh and retry on 401.
- **Strong Graph integration** with support for GCC cloud and empty-body responses.
- **Transport parity** with AWS MCP (`streamable-http`).

---

Feel free to extend this document with sequence diagrams or flowcharts if desired. For any updates to configuration or additional Graph operations, keep this architecture file in sync.


