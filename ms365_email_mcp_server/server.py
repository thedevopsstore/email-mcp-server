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

# Transport configuration (matching AWS API MCP server pattern)
# Reference: https://github.com/awslabs/mcp/blob/main/src/aws-api-mcp-server/awslabs/aws_api_mcp_server/core/common/config.py#L67
def get_transport_from_env() -> tuple[str, str]:
    """
    Get transport value from environment variable, with validation.
    Returns (env_value, fastmcp_value) tuple.
    - env_value: 'stdio' or 'streamable-http' (matches AWS API MCP server)
    - fastmcp_value: 'stdio' or 'http' (what FastMCP expects)
    """
    transport = os.getenv("TRANSPORT", "streamable-http").lower()
    if transport not in ["stdio", "streamable-http"]:
        raise ValueError(f"Invalid transport: {transport}. Must be 'stdio' or 'streamable-http'")
    
    # Map 'streamable-http' to 'http' for FastMCP
    fastmcp_transport = "http" if transport == "streamable-http" else transport
    return transport, fastmcp_transport

TRANSPORT_ENV, TRANSPORT = get_transport_from_env()

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
        self._access_token: Optional[str] = None
    
    def get_access_token(self) -> str:
        """
        Get access token using Client Credentials Flow.
        
        Reference: https://learn.microsoft.com/en-us/graph/auth-v2-service?tabs=http#step-3-request-an-access-token
        """
        if self._access_token:
            return self._access_token
        
        # Acquire token for client (app-only authentication)
        result = self.app.acquire_token_for_client(scopes=self.scope)
        
        if "access_token" in result:
            self._access_token = result["access_token"]
            logger.info("Access token acquired successfully")
            return self._access_token
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
        self, method: str, endpoint: str, **kwargs
    ) -> Any:
        """
        Make authenticated request to Microsoft Graph API.
        
        Reference: https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http
        """
        token = self.get_access_token()
        
        # Replace /me/ with user-specific endpoint if user_identifier is set
        if self.user_identifier and endpoint.startswith("me/"):
            endpoint = endpoint.replace("me/", f"users/{self.user_identifier}/", 1)
        elif self.user_identifier and "/me/" in endpoint:
            endpoint = endpoint.replace("/me/", f"/users/{self.user_identifier}/")
        
        # Ensure endpoint starts with /v1.0 or /beta
        if not endpoint.startswith("/v1.0") and not endpoint.startswith("/beta"):
            endpoint = f"/v1.0/{endpoint.lstrip('/')}"
        
        url = f"{self.graph_base}{endpoint}"
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        async with httpx.AsyncClient() as client:
            response = await client.request(method, url, headers=headers, **kwargs)
            response.raise_for_status()
            return response.json()
    
    async def list_mail_messages(
        self, folder_id: Optional[str] = None, top: int = 25
    ) -> list:
        """
        List mail messages from inbox or a specific folder.
        
        Reference: https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=http
        """
        if folder_id:
            endpoint = f"me/mailFolders/{folder_id}/messages"
        else:
            endpoint = "me/messages"
        
        params = {
            "$top": top,
            "$orderby": "receivedDateTime desc"
        }
        
        result = await self._make_request("GET", endpoint, params=params)
        return result.get("value", [])
    
    async def list_mail_folders(self) -> list:
        """List all mail folders."""
        result = await self._make_request("GET", "me/mailFolders")
        return result.get("value", [])
    
    async def list_mail_folder_messages(self, folder_id: str, top: int = 25) -> list:
        """List messages from a specific folder."""
        return await self.list_mail_messages(folder_id=folder_id, top=top)
    
    async def get_mail_message(self, message_id: str) -> dict:
        """Get a specific mail message by ID."""
        return await self._make_request("GET", f"me/messages/{message_id}")
    
    async def send_mail(
        self, to: str, subject: str, body: str, body_type: str = "HTML"
    ) -> dict:
        """
        Send an email.
        
        Reference: https://learn.microsoft.com/en-us/graph/api/user-post-messages?view=graph-rest-1.0&tabs=http
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
        return await self._make_request("POST", "me/sendMail", json=payload)
    
    async def delete_mail_message(self, message_id: str) -> None:
        """Delete a mail message."""
        await self._make_request("DELETE", f"me/messages/{message_id}")
    
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
    description="List mail messages from inbox or a specific folder. Returns a list of messages with their details including subject, sender, received date, and message ID.",
    annotations=ToolAnnotations(
        title="List mail messages",
        readOnlyHint=True,
        openWorldHint=False,
    ),
)
async def list_mail_messages(
    folder_id: Annotated[
        Optional[str],
        Field(description="Optional folder ID. If not provided, lists from inbox.")
    ] = None,
    top: Annotated[
        int,
        Field(description="Number of messages to retrieve (default: 25)", ge=1, le=100)
    ] = 25,
    ctx: Context = None,
) -> dict[str, Any]:
    """List mail messages from inbox or a specific folder."""
    try:
        client = get_client()
        messages = await client.list_mail_messages(folder_id=folder_id, top=top)
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
    description="List messages from a specific folder by folder ID. Returns messages with their details.",
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
    ctx: Context = None,
) -> dict[str, Any]:
    """List messages from a specific folder."""
    try:
        client = get_client()
        messages = await client.list_mail_folder_messages(folder_id=folder_id, top=top)
        return {"messages": messages, "count": len(messages)}
    except Exception as e:
        error_message = f"Error listing folder messages: {str(e)}"
        logger.error(error_message)
        if ctx:
            await ctx.error(error_message)
        raise


@server.tool(
    name="get-mail-message",
    description="Get a specific mail message by its ID. Returns full message details including body, attachments, and metadata.",
    annotations=ToolAnnotations(
        title="Get mail message",
        readOnlyHint=True,
        openWorldHint=False,
    ),
)
async def get_mail_message(
    message_id: Annotated[
        str,
        Field(description="Message ID to retrieve")
    ],
    ctx: Context = None,
) -> dict[str, Any]:
    """Get a specific mail message by ID."""
    try:
        client = get_client()
        message = await client.get_mail_message(message_id)
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
        await client.delete_mail_message(message_id)
        return {"success": True, "message": "Message deleted successfully"}
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
    logger.info(f"Transport: {TRANSPORT_ENV}")
    logger.info(f"Stateless HTTP: {STATELESS_HTTP}")
    
    # Run the server with explicit transport
    # TRANSPORT is mapped to FastMCP's expected values ('stdio' or 'http')
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
