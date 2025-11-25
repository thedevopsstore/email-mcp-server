# Email Operations via MCP Server

**Use the MS365 Email MCP Server for ALL email operations.**

## Tools

**Reading**: `list-mail-messages` (default: unread Inbox only; set `unread_only=false` for all), `list-mail-folders` (get folder IDs), `list-mail-folder-messages` (specific folder), `get-mail-message` (full content by ID).

**Sending**: `send-mail` (to, subject, body), `create-draft-email` (draft).

**Managing**: `delete-mail-message` (by ID), `move-mail-message` (message ID + folder ID).

## Default Behavior

`list-mail-messages` returns only **unread messages from Inbox** by default (minimizes tokens). Set `unread_only=false` for all messages. For other folders, use `list-mail-folders` first to get folder IDs.

## Workflow

1. List messages → 2. Get full content with `get-mail-message` if needed → 3. Act using message IDs.

