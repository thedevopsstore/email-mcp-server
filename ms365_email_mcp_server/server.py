#!/usr/bin/env python3
"""
MS365 Email MCP Server
A Model Context Protocol server for Microsoft 365 Outlook email operations.
Uses OAuth 2.0 Client Credentials Flow for authentication.
Built with FastMCP for simplified server implementation.
"""
import os
import sys
from typing import Annotated, Any, Optional
from fastmcp import Context, FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field
from msal import ConfidentialClientApplication
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

# Transport configuration - hardcoded to streamable-http
# Reference: https://github.com/awslabs/mcp/blob/main/src/aws-api-mcp-server/awslabs/aws_api_mcp_server/core/common/config.py#L67
TRANSPORT = "streamable-http"  # FastMCP expects 'streamable-http' for HTTP/SSE transport

# Initialize FastMCP server
server = FastMCP(
    name="MS365-Email-MCP",
    log_level=log_level,
    host=HOST,
    port=PORT,
    stateless_http=STATELESS_HTTP,
)


class MS365EmailClient:
    """
    Microsoft 365 Email API client using Client Credentials Flow.
    
    Reference: https://learn.microsoft.com/en-us/graph/auth-v2-service?tabs=http
    
    For shared mailboxes or app-only authentication, use user_identifier
    (UserPrincipalName or Graph ID) instead of /me/ endpoints.
    """
    
    def __init__(self, user_identifier: Optional[str] = None):
        self.client_id = os.getenv("MS365_CLIENT_ID")
        self.client_secret = os.getenv("MS365_CLIENT_SECRET")
        self.tenant_id = os.getenv("MS365_TENANT_ID")
        # User identifier for shared mailboxes (UserPrincipalName or Graph ID)
        # If not provided, defaults to /me/ (requires delegated permissions)
        self.user_identifier = user_identifier or os.getenv("MS365_USER_IDENTIFIER")
        
        # Determine cloud type (commercial or gov)
        cloud_type = os.getenv("MS365_CLOUD_TYPE", "commercial").lower()
        if cloud_type in ["gov", "government", "usgov"]:
            self.authority_base = "https://login.microsoftonline.us"
            self.graph_base = "https://graph.microsoft.us"
        else:
            self.authority_base = "https://login.microsoftonline.com"
            self.graph_base = "https://graph.microsoft.com"
        
        if not all([self.client_id, self.client_secret, self.tenant_id]):
            raise ValueError(
                "MS365_CLIENT_ID, MS365_CLIENT_SECRET, and MS365_TENANT_ID must be set"
            )
        
        # Configure MSAL ConfidentialClientApplication for client credentials flow
        # Reference: https://learn.microsoft.com/en-us/graph/auth-v2-service?tabs=http
        self.authority = f"{self.authority_base}/{self.tenant_id}"
        self.scope = [f"{self.graph_base}/.default"]
        
        self.app = ConfidentialClientApplication(
            self.client_id,
            authority=self.authority,
            client_credential=self.client_secret
        )
    
    def get_access_token(self) -> str:
        """
        Get access token using Client Credentials Flow.
        
        Reference: https://learn.microsoft.com/en-us/graph/auth-v2-service?tabs=http#step-3-request-an-access-token
        """
        # MSAL caches client credential tokens inside the application instance.
        # Try silent acquisition first before requesting a new token.
        result = self.app.acquire_token_silent(self.scope, account=None)

        if not result:
            result = self.app.acquire_token_for_client(scopes=self.scope)

        if "access_token" in result:
            logger.info("Access token acquired successfully (MSAL cache hit=%s)", bool(result.get("token_source") == "cache"))
            return result["access_token"]
        else:
            error = result.get("error_description", result.get("error", "Unknown error"))
            logger.error(f"Failed to acquire token: {error}")
            raise Exception(f"Failed to acquire token: {error}")
    
    def _get_user_prefix(self) -> str:
        """Get the user prefix for endpoints (/me/ or /users/{id}/)."""
        if self.user_identifier:
            return f"users/{self.user_identifier}"
        return "me"
    
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
        # Replace /me/ with user-specific endpoint if user_identifier is set
        if self.user_identifier and endpoint.startswith("me/"):
            endpoint = endpoint.replace("me/", f"users/{self.user_identifier}/", 1)
        elif self.user_identifier and "/me/" in endpoint:
            endpoint = endpoint.replace("/me/", f"/users/{self.user_identifier}/")

        # Ensure endpoint starts with /v1.0 or /beta
        if not endpoint.startswith("/v1.0") and not endpoint.startswith("/beta"):
            endpoint = f"/v1.0/{endpoint.lstrip('/')}"

        url = f"{self.graph_base}{endpoint}"

        last_response: Optional[httpx.Response] = None

        for attempt in range(2):
            token = self.get_access_token()

            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }

            async with httpx.AsyncClient() as client:
                response = await client.request(method, url, headers=headers, **kwargs)
            last_response = response

            if response.status_code == 401 and attempt == 0:
                logger.warning("Access token expired or invalid. Refreshing token and retrying once...")
                if self.app.token_cache:
                    self.app.token_cache.clear()
                continue

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

        # If we exhausted retries, raise the last response error
        if last_response is not None:
            last_response.raise_for_status()

        raise Exception("Request failed without a valid HTTP response")
    
    async def list_mail_messages(
        self, folder_id: Optional[str] = None, top: int = 25, unread_only: bool = True
    ) -> list:
        """
        List mail messages from inbox or a specific folder.
        
        By default, only lists unread messages from the Inbox folder to avoid scanning
        all folders (inbox, sent items, deleted items, etc.) and minimize token usage.
        
        Reference: https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http
        """
        # Default to Inbox folder to avoid scanning all folders
        # Use well-known folder name "Inbox" which is supported by Microsoft Graph
        if folder_id:
            endpoint = f"me/mailFolders/{folder_id}/messages"
        else:
            endpoint = "me/mailFolders/Inbox/messages"
        
        # Use $select to reduce response size and improve performance
        # Only fetch essential properties to minimize token usage
        params = {
            "$top": top,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,sender,receivedDateTime,isRead,hasAttachments,bodyPreview"
        }
        
        # By default, filter to only unread messages to minimize token usage
        if unread_only:
            params["$filter"] = "isRead eq false"
        
        result = await self._make_request("GET", endpoint, params=params)
        return result.get("value", [])
    
    async def list_mail_folders(self) -> list:
        """List all mail folders."""
        result = await self._make_request("GET", "me/mailFolders")
        return result.get("value", [])
    
    async def list_mail_folder_messages(self, folder_id: str, top: int = 25, unread_only: bool = True) -> list:
        """List messages from a specific folder. By default, only returns unread messages."""
        return await self.list_mail_messages(folder_id=folder_id, top=top, unread_only=unread_only)
    
    async def get_mail_message(self, message_id: str, mark_as_read: bool = True) -> dict:
        """
        Get a specific mail message by ID.
        
        By default, marks the message as read after retrieving it.
        
        Reference: https://learn.microsoft.com/en-us/graph/api/message-get?view=graph-rest-1.0&tabs=http
        """
        message = await self._make_request("GET", f"me/messages/{message_id}")
        
        # Automatically mark message as read if requested (default behavior)
        if mark_as_read:
            try:
                await self.mark_message_as_read(message_id)
            except Exception as e:
                logger.warning(f"Failed to mark message {message_id} as read: {e}")
                # Continue even if marking as read fails
        
        return message
    
    async def mark_message_as_read(self, message_id: str) -> dict:
        """
        Mark a message as read by updating the isRead property.
        
        Reference: https://learn.microsoft.com/en-us/graph/api/message-update?view=graph-rest-1.0&tabs=http
        """
        payload = {"isRead": True}
        return await self._make_request("PATCH", f"me/messages/{message_id}", json=payload)
    
    async def mark_message_as_unread(self, message_id: str) -> dict:
        """
        Mark a message as unread by updating the isRead property.
        
        Reference: https://learn.microsoft.com/en-us/graph/api/message-update?view=graph-rest-1.0&tabs=http
        """
        payload = {"isRead": False}
        return await self._make_request("PATCH", f"me/messages/{message_id}", json=payload)
    
    async def send_mail(
        self, to: str, subject: str, body: str, body_type: str = "HTML"
    ) -> dict:
        """
        Send an email.
        
        Reference: https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http
        
        Returns 202 Accepted with empty body - the message is queued for delivery.
        """
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
        return await self._make_request("POST", "me/sendMail", return_json=False, json=payload)
    
    async def delete_mail_message(self, message_id: str) -> dict:
        """Delete a mail message. Returns 204 No Content with empty body."""
        # DELETE returns 204 No Content with no response body
        return await self._make_request("DELETE", f"me/messages/{message_id}", return_json=False)
    
    async def create_draft_email(
        self, to: str, subject: str, body: str, body_type: str = "HTML"
    ) -> dict:
        """Create a draft email."""
        payload = {
            "subject": subject,
            "body": {
                "contentType": body_type,
                "content": body
            },
            "toRecipients": [{"emailAddress": {"address": to}}]
        }
        return await self._make_request("POST", "me/messages", json=payload)
    
    async def move_mail_message(self, message_id: str, destination_id: str) -> dict:
        """Move a mail message to another folder."""
        payload = {"destinationId": destination_id}
        return await self._make_request(
            "POST", f"me/messages/{message_id}/move", json=payload
        )


# Initialize client (lazy initialization)
_ms365_client: Optional[MS365EmailClient] = None


def get_client(user_identifier: Optional[str] = None) -> MS365EmailClient:
    """
    Get or create MS365 email client.
    
    Args:
        user_identifier: Optional UserPrincipalName or Graph ID for shared mailboxes.
                        If not provided, uses MS365_USER_IDENTIFIER env var or /me/ endpoints.
    """
    global _ms365_client
    # Use provided user_identifier or environment variable
    effective_user_id = user_identifier or os.getenv("MS365_USER_IDENTIFIER")
    
    # Create new client if user_identifier changed or client doesn't exist
    if _ms365_client is None or _ms365_client.user_identifier != effective_user_id:
        _ms365_client = MS365EmailClient(user_identifier=effective_user_id)
    return _ms365_client


@server.tool(
    name="list-mail-messages",
    description="List mail messages from inbox or a specific folder. By default, only lists unread messages from the Inbox folder (not sent items or other folders) to minimize token usage. Returns a list of messages with their details including subject, sender, received date, and message ID. NOTE: This only returns previews - to actually read an email and mark it as read, you MUST use get-mail-message with the message ID.",
    annotations=ToolAnnotations(
        title="List mail messages",
        readOnlyHint=True,
        openWorldHint=False,
    ),
)
async def list_mail_messages(
    folder_id: Annotated[
        Optional[str],
        Field(description="Optional folder ID or well-known folder name (e.g., 'Inbox', 'SentItems', 'Drafts'). If not provided, defaults to Inbox folder only.")
    ] = None,
    top: Annotated[
        int,
        Field(description="Number of messages to retrieve (default: 25)", ge=1, le=100)
    ] = 25,
    unread_only: Annotated[
        bool,
        Field(description="If true, only return unread messages. If false, return all messages (read and unread). Default: true.")
    ] = True,
    ctx: Context = None,
) -> dict[str, Any]:
    """List mail messages from inbox or a specific folder. Defaults to unread messages from Inbox only to avoid scanning all folders."""
    try:
        client = get_client()
        messages = await client.list_mail_messages(folder_id=folder_id, top=top, unread_only=unread_only)
        return {"messages": messages, "count": len(messages)}
    except Exception as e:
        error_message = f"Error listing mail messages: {str(e)}"
        logger.error(error_message)
        if ctx:
            await ctx.error(error_message)
        raise


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
    ctx: Context = None,
) -> dict[str, Any]:
    """List all mail folders."""
    try:
        client = get_client()
        folders = await client.list_mail_folders()
        return {"folders": folders, "count": len(folders)}
    except Exception as e:
        error_message = f"Error listing mail folders: {str(e)}"
        logger.error(error_message)
        if ctx:
            await ctx.error(error_message)
        raise


@server.tool(
    name="list-mail-folder-messages",
    description="List messages from a specific folder by folder ID. By default, only returns unread messages to minimize token usage. Returns messages with their details. NOTE: This only returns previews - to actually read an email and mark it as read, you MUST use get-mail-message with the message ID.",
    annotations=ToolAnnotations(
        title="List folder messages",
        readOnlyHint=True,
        openWorldHint=False,
    ),
)
async def list_mail_folder_messages(
    folder_id: Annotated[
        str,
        Field(description="Folder ID to list messages from")
    ],
    top: Annotated[
        int,
        Field(description="Number of messages to retrieve (default: 25)", ge=1, le=100)
    ] = 25,
    unread_only: Annotated[
        bool,
        Field(description="If true, only return unread messages. If false, return all messages (read and unread). Default: true.")
    ] = True,
    ctx: Context = None,
) -> dict[str, Any]:
    """List messages from a specific folder. By default, only returns unread messages."""
    try:
        client = get_client()
        messages = await client.list_mail_folder_messages(folder_id=folder_id, top=top, unread_only=unread_only)
        return {"messages": messages, "count": len(messages)}
    except Exception as e:
        error_message = f"Error listing folder messages: {str(e)}"
        logger.error(error_message)
        if ctx:
            await ctx.error(error_message)
        raise


@server.tool(
    name="get-mail-message",
    description="Get a specific mail message by its ID. REQUIRED to actually read an email - automatically marks the message as read after retrieval. Returns full message details including body, attachments, and metadata. Use this (not list-mail-messages) when you need to read, process, or act on an email.",
    annotations=ToolAnnotations(
        title="Get mail message",
        readOnlyHint=True,
        openWorldHint=False,
    ),
)
async def get_mail_message(
    message_id: Annotated[
        str,
        Field(description="Message ID to retrieve. The message will be automatically marked as read. Use this to actually read an email, not just list-mail-messages.")
    ],
    ctx: Context = None,
) -> dict[str, Any]:
    """Get a specific mail message by ID. Automatically marks the message as read."""
    try:
        client = get_client()
        message = await client.get_mail_message(message_id, mark_as_read=True)
        return {"message": message}
    except Exception as e:
        error_message = f"Error getting mail message: {str(e)}"
        logger.error(error_message)
        if ctx:
            await ctx.error(error_message)
        raise


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
    user_identifier: Annotated[
        Optional[str],
        Field(description="Optional: UserPrincipalName or Graph ID for shared mailboxes. If not provided, uses MS365_USER_IDENTIFIER env var or /me/ endpoint.")
    ] = None,
    ctx: Context = None,
) -> dict[str, Any]:
    """Send an email."""
    try:
        if body_type not in ["HTML", "Text"]:
            raise ValueError("body_type must be 'HTML' or 'Text'")
        
        client = get_client(user_identifier=user_identifier)
        result = await client.send_mail(to, subject, body, body_type)
        return {"success": True, "result": result}
    except Exception as e:
        error_message = f"Error sending mail: {str(e)}"
        logger.error(error_message)
        if ctx:
            await ctx.error(error_message)
        raise


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
    ctx: Context = None,
) -> dict[str, Any]:
    """Delete a mail message."""
    try:
        client = get_client()
        result = await client.delete_mail_message(message_id)
        return {"success": True, "message": "Message deleted successfully", "result": result}
    except Exception as e:
        error_message = f"Error deleting mail message: {str(e)}"
        logger.error(error_message)
        if ctx:
            await ctx.error(error_message)
        raise


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
    ctx: Context = None,
) -> dict[str, Any]:
    """Create a draft email."""
    try:
        if body_type not in ["HTML", "Text"]:
            raise ValueError("body_type must be 'HTML' or 'Text'")
        
        client = get_client()
        draft = await client.create_draft_email(to, subject, body, body_type)
        return {"success": True, "draft": draft}
    except Exception as e:
        error_message = f"Error creating draft email: {str(e)}"
        logger.error(error_message)
        if ctx:
            await ctx.error(error_message)
        raise


@server.tool(
    name="move-mail-message",
    description="Move a mail message to another folder by specifying the message ID and destination folder ID.",
    annotations=ToolAnnotations(
        title="Move mail message",
        readOnlyHint=False,
        destructiveHint=False,
        openWorldHint=False,
    ),
)
async def move_mail_message(
    message_id: Annotated[
        str,
        Field(description="Message ID to move")
    ],
    destination_id: Annotated[
        str,
        Field(description="Destination folder ID")
    ],
    ctx: Context = None,
) -> dict[str, Any]:
    """Move a mail message to another folder."""
    try:
        client = get_client()
        result = await client.move_mail_message(message_id, destination_id)
        return {"success": True, "result": result}
    except Exception as e:
        error_message = f"Error moving mail message: {str(e)}"
        logger.error(error_message)
        if ctx:
            await ctx.error(error_message)
        raise


@server.tool(
    name="mark-mail-message-read",
    description="Mark a mail message as read by its ID. Useful for explicitly marking messages as read without retrieving the full content.",
    annotations=ToolAnnotations(
        title="Mark message as read",
        readOnlyHint=False,
        destructiveHint=False,
        openWorldHint=False,
    ),
)
async def mark_mail_message_read(
    message_id: Annotated[
        str,
        Field(description="Message ID to mark as read")
    ],
    ctx: Context = None,
) -> dict[str, Any]:
    """Mark a mail message as read."""
    try:
        client = get_client()
        result = await client.mark_message_as_read(message_id)
        return {"success": True, "message": "Message marked as read", "result": result}
    except Exception as e:
        error_message = f"Error marking message as read: {str(e)}"
        logger.error(error_message)
        if ctx:
            await ctx.error(error_message)
        raise


@server.tool(
    name="mark-mail-message-unread",
    description="Mark a mail message as unread by its ID. Useful for flagging messages that need attention later.",
    annotations=ToolAnnotations(
        title="Mark message as unread",
        readOnlyHint=False,
        destructiveHint=False,
        openWorldHint=False,
    ),
)
async def mark_mail_message_unread(
    message_id: Annotated[
        str,
        Field(description="Message ID to mark as unread")
    ],
    ctx: Context = None,
) -> dict[str, Any]:
    """Mark a mail message as unread."""
    try:
        client = get_client()
        result = await client.mark_message_as_unread(message_id)
        return {"success": True, "message": "Message marked as unread", "result": result}
    except Exception as e:
        error_message = f"Error marking message as unread: {str(e)}"
        logger.error(error_message)
        if ctx:
            await ctx.error(error_message)
        raise


def main():
    """Main entry point for the MS365 Email MCP server."""
    # Validate required environment variables
    if not all([
        os.getenv("MS365_CLIENT_ID"),
        os.getenv("MS365_CLIENT_SECRET"),
        os.getenv("MS365_TENANT_ID")
    ]):
        error_message = (
            "MS365_CLIENT_ID, MS365_CLIENT_SECRET, and MS365_TENANT_ID must be set"
        )
        logger.error(error_message)
        raise ValueError(error_message)
    
    logger.info(f"Starting MS365 Email MCP Server on {HOST}:{PORT}")
    logger.info(f"Transport: {TRANSPORT}")
    logger.info(f"Stateless HTTP: {STATELESS_HTTP}")
    
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
