# Email Operations via MCP Server

**Use the MS365 Email MCP Server for ALL email operations.**

## Tools

**Reading**: `list-mail-messages` (default: unread Inbox only; set `unread_only=false` for all), `list-mail-folders` (get folder IDs), `list-mail-folder-messages` (specific folder), `get-mail-message` (full content by ID - **automatically marks as read**).

**Sending**: `send-mail` (to, subject, body), `create-draft-email` (draft).

**Managing**: `delete-mail-message` (by ID), `move-mail-message` (message ID + folder ID), `mark-mail-message-read`, `mark-mail-message-unread`.

## Default Behavior

`list-mail-messages` returns only **unread messages from Inbox** by default (minimizes tokens). Set `unread_only=false` for all messages. For other folders, use `list-mail-folders` first to get folder IDs.

## IMPORTANT: Reading Emails

**`list-mail-messages` only shows previews and does NOT mark emails as read.** To actually read an email and mark it as read, you MUST use `get-mail-message` with the message ID. Always use `get-mail-message` when you need to:
- Read the full email content
- Process or summarize an email
- Reply to an email
- Perform any action that requires reading the email

## Workflow

1. Use `list-mail-messages` to scan/see what emails exist
2. **MUST use `get-mail-message` with message ID to read any email** (marks as read automatically)
3. Act using message IDs (send, delete, move, etc.)

