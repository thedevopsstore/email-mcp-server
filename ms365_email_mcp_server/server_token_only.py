#!/usr/bin/env python3
"""
MS365 Email MCP Server (Token-Only Version)
A Model Context Protocol server for Microsoft 365 Outlook email operations.
Uses token-based authentication - agents must provide JWT tokens for Microsoft Graph API.
Built with FastMCP for simplified server implementation.
"""
import os
import sys
from typing import Annotated, Any, Optional, Dict
from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field
from fastmcp.server.dependencies import get_http_headers
import httpx
from loguru import logger

# Configure logging
logger.remove()
log_level = os.getenv("LOG_LEVEL", "INFO").upper()
logger.add(sys.stderr, level=log_level)

# Server configuration
HOST = os.getenv("HOST", "0.0.0.0")
PORT = int(os.getenv("PORT", "8100"))
STATELESS_HTTP = os.getenv("STATELESS_HTTP", "true").lower() == "true"

TRANSPORT = "streamable-http"  # FastMCP expects 'streamable-http' for HTTP/SSE transport

# Initialize FastMCP server
server = FastMCP(
    name="MS365-Email-MCP",
    log_level=log_level,
    host=HOST,
    port=PORT,
    stateless_http=STATELESS_HTTP,
)


def extract_request_auth_from_headers() -> tuple[Optional[str], Optional[str]]:
    """
    Extract request authentication context from HTTP headers:
    - access_token from `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization: Bearer <token>`
    - user_identifier from `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier`
    
    Uses FastMCP's built-in get_http_headers() which automatically handles
    request context and works even when MCP session isn't fully established.
    Never raises exceptions - returns empty dict if no request context.

    Expected header:
    - X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization: Bearer <token>
    - X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier: <UPN or Graph user id>
    """
    # FastMCP's get_http_headers() never raises exceptions
    # Returns empty dict if no request context
    headers = get_http_headers(include_all=True)

    access_token: Optional[str] = None
    auth_header = "x-amzn-bedrock-agentcore-runtime-custom-ms365-authorization"
    auth = headers.get(auth_header)
    if isinstance(auth, str) and auth.strip():
        auth = auth.strip()
        # Accept either "Bearer <token>" or raw token
        if auth.lower().startswith("bearer "):
            access_token = auth[7:].strip() or None
        else:
            access_token = auth or None

    user_identifier: Optional[str] = None
    user_header = "x-amzn-bedrock-agentcore-runtime-custom-ms365-useridentifier"
    raw_user = headers.get(user_header)
    if isinstance(raw_user, str) and raw_user.strip():
        user_identifier = raw_user.strip()

    return access_token, user_identifier


class MS365EmailClient:
    """
    Microsoft 365 Email API client using token-based authentication.
    
    Agents must provide a JWT token for Microsoft Graph API. The token is used
    directly for API calls - no token exchange or refresh is performed.
    
    Reference: https://learn.microsoft.com/en-us/graph/auth-v2-service?tabs=http
    
    For shared mailboxes, use user_identifier (UserPrincipalName or Graph ID)
    instead of /me/ endpoints.
    """
    
    def __init__(
        self, 
        access_token: str,
        user_identifier: Optional[str] = None,
        cloud_type: Optional[str] = None
    ):
        if not access_token:
            raise ValueError("access_token is required for token-based authentication")
        
        # Remove "Bearer " prefix if present
        if access_token.startswith("Bearer "):
            self.access_token = access_token[7:]
        else:
            self.access_token = access_token
        
        # User identifier for shared mailboxes (UserPrincipalName or Graph ID)
        # If not provided, defaults to /me/ (requires delegated permissions)
        self.user_identifier = user_identifier or os.getenv("MS365_USER_IDENTIFIER")
        
        # Determine cloud type (commercial or gov)
        effective_cloud_type = (cloud_type or os.getenv("MS365_CLOUD_TYPE", "commercial")).lower()
        
        if effective_cloud_type in ["gov", "government", "usgov"]:
            self.graph_base = "https://graph.microsoft.us"
        else:
            self.graph_base = "https://graph.microsoft.com"
        
        logger.debug("Using token-based authentication mode")
    
    def _build_endpoint(self, endpoint: str) -> str:
        """
        Build endpoint with correct user prefix.
        
        Args:
            endpoint: Endpoint starting with "me/" (e.g., "me/messages")
        
        Returns:
            Endpoint with "me/" or "users/{id}/" prefix based on user_identifier
        """
        if self.user_identifier:
            return endpoint.replace("me/", f"users/{self.user_identifier}/", 1)
        return endpoint
    
    async def _make_request(
        self, method: str, endpoint: str, return_json: bool = True, **kwargs
    ) -> Any:
        """
        Make authenticated request to Microsoft Graph API.
        
        Reference: https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http
        
        Args:
            method: HTTP method (GET, POST, DELETE, etc.)
            endpoint: API endpoint (e.g., "me/sendMail", "me/messages")
            return_json: Whether to parse JSON response (default: True)
                        Set to False for endpoints that return empty body (e.g., sendMail returns 202)
        """
        # Note: Endpoint should already be transformed by _build_endpoint() in calling methods
        # Ensure endpoint starts with /v1.0 or /beta
        if not endpoint.startswith(("/v1.0", "/beta")):
            endpoint = f"/v1.0/{endpoint.lstrip('/')}"

        url = f"{self.graph_base}{endpoint}"

        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

        async with httpx.AsyncClient() as client:
            response = await client.request(method, url, headers=headers, **kwargs)

        # If token is expired/invalid, log warning before raising
        if response.status_code == 401:
            logger.warning("Access token expired or invalid. Agent should refresh the token.")
        
        response.raise_for_status()

        # Some endpoints (like sendMail) return 202 Accepted with empty body
        # Reference: https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http
        if not return_json or response.status_code in (202, 204):
            return {"status": response.status_code, "status_text": response.reason_phrase}

        # Try to parse JSON, but handle empty responses gracefully
        text = response.text.strip()
        if not text:
            return {"status": response.status_code, "status_text": response.reason_phrase}

        return response.json()
    
    async def list_mail_messages(
        self, 
        folder_id: Optional[str] = None, 
        top: int = 25, 
        unread_only: bool = True
    ) -> list:
        """
        List mail messages from inbox or a specific folder.
        
        Note: This endpoint only returns bodyPreview (first 255 characters) per Microsoft Graph API.
        To get full email body content, use get_mail_message() with the message ID.
        
        Note: This endpoint does NOT support marking messages as read. To mark a message as read,
        use get_mail_message() which automatically marks messages as read, or use mark_message_as_read().
        
        By default, only lists unread messages from the Inbox folder to avoid scanning
        all folders (inbox, sent items, deleted items, etc.) and minimize token usage.
        
        Reference: https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http
        
        Args:
            folder_id: Optional folder ID. If not provided, defaults to Inbox.
            top: Number of messages to retrieve (default: 25)
            unread_only: If True, only returns unread messages (default: True)
        """
        # Default to Inbox folder to avoid scanning all folders
        # Use well-known folder name "Inbox" which is supported by Microsoft Graph
        if folder_id:
            endpoint = self._build_endpoint(f"me/mailFolders/{folder_id}/messages")
        else:
            endpoint = self._build_endpoint("me/mailFolders/Inbox/messages")
        
        # Use $select to reduce response size and improve performance
        # Note: List messages endpoint only returns bodyPreview, not full body
        # To get full body, use get_mail_message() with the message ID
        params = {
            "$top": top,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,sender,receivedDateTime,isRead,hasAttachments,bodyPreview"
        }
        
        # By default, filter to only unread messages to minimize token usage
        if unread_only:
            params["$filter"] = "isRead eq false"
        
        result = await self._make_request("GET", endpoint, params=params)
        messages = result.get("value", [])
        
        return messages
    
    async def list_mail_folders(self) -> list:
        """List all mail folders."""
        endpoint = self._build_endpoint("me/mailFolders")
        result = await self._make_request("GET", endpoint)
        return result.get("value", [])
    
    async def get_mail_message(self, message_id: str, mark_as_read: bool = True) -> dict:
        """
        Get a specific mail message by ID.
        
        Args:
            message_id: Message ID to retrieve
            mark_as_read: If True, automatically marks the message as read (default: True)
        
        Reference: https://learn.microsoft.com/en-us/graph/api/message-get?view=graph-rest-1.0&tabs=http
        """
        endpoint = self._build_endpoint(f"me/messages/{message_id}")
        message = await self._make_request("GET", endpoint)
        
        # Automatically mark message as read when retrieved
        if mark_as_read:
            await self.mark_message_as_read(message_id)
        
        return message
    
    async def mark_message_as_read(self, message_id: str) -> dict:
        """
        Mark a message as read by updating the isRead property.
        
        Reference: https://learn.microsoft.com/en-us/graph/api/message-update?view=graph-rest-1.0&tabs=http
        """
        payload = {"isRead": True}
        endpoint = self._build_endpoint(f"me/messages/{message_id}")
        return await self._make_request("PATCH", endpoint, json=payload)
    
    async def mark_message_as_unread(self, message_id: str) -> dict:
        """
        Mark a message as unread by updating the isRead property.
        
        Reference: https://learn.microsoft.com/en-us/graph/api/message-update?view=graph-rest-1.0&tabs=http
        """
        payload = {"isRead": False}
        endpoint = self._build_endpoint(f"me/messages/{message_id}")
        return await self._make_request("PATCH", endpoint, json=payload)
    
    async def send_mail(
        self, to: str, subject: str, body: str, body_type: str = "HTML"
    ) -> dict:
        """
        Send an email.
        
        Reference: https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http
        
        Returns 202 Accepted with empty body - the message is queued for delivery.
        
        Args:
            body_type: Must be "HTML" or "Text" (default: "HTML")
        """
        if body_type not in ["HTML", "Text"]:
            raise ValueError("body_type must be 'HTML' or 'Text'")
        
        payload = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": body_type,
                    "content": body
                },
                "toRecipients": [{"emailAddress": {"address": to}}]
            },
            "saveToSentItems": "true"
        }
        # sendMail returns 202 Accepted with no response body
        endpoint = self._build_endpoint("me/sendMail")
        return await self._make_request("POST", endpoint, return_json=False, json=payload)
    
    async def delete_mail_message(self, message_id: str) -> dict:
        """Delete a mail message. Returns 204 No Content with empty body."""
        # DELETE returns 204 No Content with no response body
        endpoint = self._build_endpoint(f"me/messages/{message_id}")
        return await self._make_request("DELETE", endpoint, return_json=False)
    
    async def create_draft_email(
        self, to: str, subject: str, body: str, body_type: str = "HTML"
    ) -> dict:
        """
        Create a draft email.
        
        Args:
            body_type: Must be "HTML" or "Text" (default: "HTML")
        """
        if body_type not in ["HTML", "Text"]:
            raise ValueError("body_type must be 'HTML' or 'Text'")
        
        payload = {
            "subject": subject,
            "body": {
                "contentType": body_type,
                "content": body
            },
            "toRecipients": [{"emailAddress": {"address": to}}]
        }
        endpoint = self._build_endpoint("me/messages")
        return await self._make_request("POST", endpoint, json=payload)

def get_client(
    access_token: Optional[str] = None,
    user_identifier: Optional[str] = None,
    cloud_type: Optional[str] = None
) -> MS365EmailClient:
    """
    Get or create MS365 email client.
    
    Token-based authentication only - access_token is required.
    
    Priority:
    1. Function parameters (access_token, user_identifier)
    2. HTTP headers:
       - X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization header for access_token
       - X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier for user_identifier
    3. Environment variable (MS365_USER_IDENTIFIER) for user_identifier
    
    Args:
        access_token: Optional JWT token for Microsoft Graph API (Bearer token, with or without "Bearer " prefix).
                     If not provided, will attempt to extract from X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization header.
        user_identifier: Optional UserPrincipalName or Graph ID for shared mailboxes.
                        If not provided, will attempt to extract from custom header or environment variable.
        cloud_type: Optional cloud type: "commercial" or "gov"
    """
    # Read headers once (only if we need them)
    hdr_token: Optional[str] = None
    hdr_user: Optional[str] = None
    if not access_token or not user_identifier:
        hdr_token, hdr_user = extract_request_auth_from_headers()

    # Priority: function parameter > headers
    effective_token = access_token or hdr_token
    if not effective_token:
        raise ValueError(
            "access_token is required. Provide it via tool parameter or "
            "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization header (Bearer <token>)"
        )
    
    # Determine effective user identifier
    # Priority: function parameter > custom header > environment variable
    effective_user_id = user_identifier or hdr_user or os.getenv("MS365_USER_IDENTIFIER")

    # NOTE: We intentionally do not cache MS365EmailClient instances by token.
    # Caching would keep bearer tokens in process memory indefinitely and provides
    # little benefit (this client is lightweight and we create httpx clients per request).
    return MS365EmailClient(
        access_token=effective_token,
        user_identifier=effective_user_id,
        cloud_type=cloud_type,
    )


@server.tool(
    name="list-mail-messages",
    description="Lists email PREVIEWS only (bodyPreview field, ~255 chars). ⚠️ WARNING: This does NOT return full email content. Use get-mail-message with the message ID to read full content. ⚠️ NOTE: This endpoint cannot mark messages as read. Use get-mail-message to read full content and automatically mark as read. By default, returns unread messages from the Inbox.",
    annotations=ToolAnnotations(
        title="List mail messages",
        readOnlyHint=True,
        openWorldHint=False,
    ),
)
async def list_mail_messages(
    access_token: Annotated[
        Optional[str],
        Field(description="Optional: JWT token for Microsoft Graph API (Bearer token). If not provided, will be extracted from the X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization header (Bearer <token>).")
    ] = None,
    folder_id: Annotated[
        Optional[str],
        Field(description="Optional folder ID. If not provided, lists from Inbox folder.")
    ] = None,
    top: Annotated[
        int,
        Field(description="Number of messages to retrieve (default: 25)", ge=1, le=100)
    ] = 25,
    unread_only: Annotated[
        bool,
        Field(description="If True, only returns unread messages. Set to False to get all messages including read ones. (default: True)")
    ] = True,
    user_identifier: Annotated[
        Optional[str],
        Field(description="Optional: UserPrincipalName or Graph ID for shared mailboxes.")
    ] = None,
) -> dict[str, Any]:
    """List mail messages from inbox or a specific folder. Returns previews only - use get-mail-message for full body content."""
    client = get_client(
        access_token=access_token,
        user_identifier=user_identifier
    )
    messages = await client.list_mail_messages(
        folder_id=folder_id, 
        top=top, 
        unread_only=unread_only
    )
    return {"messages": messages, "count": len(messages)}


@server.tool(
    name="list-mail-folders",
    description="List all mail folders in the mailbox. Returns folder names, IDs, and other metadata.",
    annotations=ToolAnnotations(
        title="List mail folders",
        readOnlyHint=True,
        openWorldHint=False,
    ),
)
async def list_mail_folders(
    access_token: Annotated[
        Optional[str],
        Field(description="Optional: JWT token for Microsoft Graph API (Bearer token). If not provided, will be extracted from the X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization header (Bearer <token>).")
    ] = None,
    user_identifier: Annotated[
        Optional[str],
        Field(description="Optional: UserPrincipalName or Graph ID for shared mailboxes.")
    ] = None,
) -> dict[str, Any]:
    """List all mail folders."""
    client = get_client(access_token=access_token, user_identifier=user_identifier)
    folders = await client.list_mail_folders()
    return {"folders": folders, "count": len(folders)}


@server.tool(
    name="get-mail-message",
    description="⚠️ REQUIRED for reading full email content. Use this after list-mail-messages to get complete email body. Automatically marks message as read.",
    annotations=ToolAnnotations(
        title="Get mail message",
        readOnlyHint=True,
        openWorldHint=False,
    ),
)
async def get_mail_message(
    message_id: Annotated[
        str,
        Field(description="Message ID to retrieve. The message will be automatically marked as read.")
    ],
    access_token: Annotated[
        Optional[str],
        Field(description="Optional: JWT token for Microsoft Graph API (Bearer token). If not provided, will be extracted from the X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization header (Bearer <token>).")
    ] = None,
    user_identifier: Annotated[
        Optional[str],
        Field(description="Optional: UserPrincipalName or Graph ID for shared mailboxes.")
    ] = None,
) -> dict[str, Any]:
    """Get a specific mail message by ID. Automatically marks the message as read."""
    client = get_client(access_token=access_token, user_identifier=user_identifier)
    message = await client.get_mail_message(message_id, mark_as_read=True)
    return {"message": message}


@server.tool(
    name="send-mail",
    description="Send an email to a recipient. The email will be sent immediately and saved to sent items. For shared mailboxes, provide user_identifier (UserPrincipalName or Graph ID).",
    annotations=ToolAnnotations(
        title="Send email",
        readOnlyHint=False,
        destructiveHint=False,
        openWorldHint=False,
    ),
)
async def send_mail(
    to: Annotated[
        str,
        Field(description="Recipient email address")
    ],
    subject: Annotated[
        str,
        Field(description="Email subject")
    ],
    body: Annotated[
        str,
        Field(description="Email body content")
    ],
    body_type: Annotated[
        str,
        Field(description="Body content type: 'HTML' or 'Text' (default: 'HTML')")
    ] = "HTML",
    access_token: Annotated[
        Optional[str],
        Field(description="Optional: JWT token for Microsoft Graph API (Bearer token). If not provided, will be extracted from the X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization header (Bearer <token>).")
    ] = None,
    user_identifier: Annotated[
        Optional[str],
        Field(description="Optional: UserPrincipalName or Graph ID for shared mailboxes.")
    ] = None,
) -> dict[str, Any]:
    """Send an email."""
    client = get_client(access_token=access_token, user_identifier=user_identifier)
    result = await client.send_mail(to, subject, body, body_type)
    return {"success": True, "result": result}


@server.tool(
    name="delete-mail-message",
    description="Delete a mail message by its ID. This action cannot be undone.",
    annotations=ToolAnnotations(
        title="Delete mail message",
        readOnlyHint=False,
        destructiveHint=True,
        openWorldHint=False,
    ),
)
async def delete_mail_message(
    message_id: Annotated[
        str,
        Field(description="Message ID to delete")
    ],
    access_token: Annotated[
        Optional[str],
        Field(description="Optional: JWT token for Microsoft Graph API (Bearer token). If not provided, will be extracted from the X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization header (Bearer <token>).")
    ] = None,
    user_identifier: Annotated[
        Optional[str],
        Field(description="Optional: UserPrincipalName or Graph ID for shared mailboxes.")
    ] = None,
) -> dict[str, Any]:
    """Delete a mail message."""
    client = get_client(access_token=access_token, user_identifier=user_identifier)
    result = await client.delete_mail_message(message_id)
    return {"success": True, "message": "Message deleted successfully", "result": result}


@server.tool(
    name="create-draft-email",
    description="Create a draft email without sending it. The draft will be saved in the drafts folder.",
    annotations=ToolAnnotations(
        title="Create draft email",
        readOnlyHint=False,
        destructiveHint=False,
        openWorldHint=False,
    ),
)
async def create_draft_email(
    to: Annotated[
        str,
        Field(description="Recipient email address")
    ],
    subject: Annotated[
        str,
        Field(description="Email subject")
    ],
    body: Annotated[
        str,
        Field(description="Email body content")
    ],
    body_type: Annotated[
        str,
        Field(description="Body content type: 'HTML' or 'Text' (default: 'HTML')")
    ] = "HTML",
    access_token: Annotated[
        Optional[str],
        Field(description="Optional: JWT token for Microsoft Graph API (Bearer token). If not provided, will be extracted from the X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization header (Bearer <token>).")
    ] = None,
    user_identifier: Annotated[
        Optional[str],
        Field(description="Optional: UserPrincipalName or Graph ID for shared mailboxes.")
    ] = None,
) -> dict[str, Any]:
    """Create a draft email."""
    client = get_client(access_token=access_token, user_identifier=user_identifier)
    draft = await client.create_draft_email(to, subject, body, body_type)
    return {"success": True, "draft": draft}


def main():
    """Main entry point for the MS365 Email MCP server."""
    logger.info(f"Starting MS365 Email MCP Server (Token-Only) on {HOST}:{PORT}")
    logger.info(f"Transport: {TRANSPORT}")
    logger.info(f"Stateless HTTP: {STATELESS_HTTP}")
    logger.info("Authentication: Token-based only (agents must provide access_token)")
    
    # Run the server with explicit transport
    # TRANSPORT: 'stdio' or 'streamable-http' (FastMCP accepts these values)
    server.run(transport=TRANSPORT)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.info("Server stopped")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Server error: {e}", exc_info=True)
        sys.exit(1)

