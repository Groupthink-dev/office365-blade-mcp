"""Shared constants, types, and write-gate for Office 365 Blade MCP server."""

from __future__ import annotations

import logging
import os

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Defaults
# ---------------------------------------------------------------------------

DEFAULT_LIMIT = 20
MAX_BATCH_SIZE = 50
MAX_BODY_CHARS = 50_000
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# ---------------------------------------------------------------------------
# Graph API field selections (token efficiency)
# ---------------------------------------------------------------------------

EMAIL_LIST_FIELDS = [
    "id",
    "conversationId",
    "parentFolderId",
    "receivedDateTime",
    "subject",
    "from",
    "toRecipients",
    "isRead",
    "flag",
    "importance",
    "hasAttachments",
    "bodyPreview",
]

EMAIL_READ_FIELDS = [
    "id",
    "conversationId",
    "parentFolderId",
    "receivedDateTime",
    "sentDateTime",
    "subject",
    "from",
    "toRecipients",
    "ccRecipients",
    "bccRecipients",
    "replyTo",
    "isRead",
    "flag",
    "importance",
    "hasAttachments",
    "internetMessageId",
    "body",
    "bodyPreview",
]

EVENT_LIST_FIELDS = [
    "id",
    "subject",
    "start",
    "end",
    "location",
    "organizer",
    "isAllDay",
    "isCancelled",
    "responseStatus",
    "showAs",
    "importance",
    "isOnlineMeeting",
    "onlineMeetingUrl",
]

EVENT_READ_FIELDS = EVENT_LIST_FIELDS + [
    "body",
    "attendees",
    "recurrence",
    "categories",
    "webLink",
    "createdDateTime",
    "lastModifiedDateTime",
]

TASK_LIST_FIELDS = [
    "id",
    "title",
    "status",
    "importance",
    "createdDateTime",
    "lastModifiedDateTime",
    "dueDateTime",
    "completedDateTime",
    "body",
]

# ---------------------------------------------------------------------------
# Permission scopes
# ---------------------------------------------------------------------------

SCOPES_READ = [
    "User.Read",
    "Mail.Read",
    "Calendars.Read",
    "Tasks.Read",
]

SCOPES_READWRITE = [
    "User.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Tasks.ReadWrite",
]

SCOPES_CLIENT_CREDENTIALS = [
    "https://graph.microsoft.com/.default",
]


def get_scopes() -> list[str]:
    """Return scopes based on write mode."""
    if is_write_enabled():
        return SCOPES_READWRITE
    return SCOPES_READ


# ---------------------------------------------------------------------------
# Write gate
# ---------------------------------------------------------------------------


def is_write_enabled() -> bool:
    """Check if write operations are enabled via env var."""
    return os.environ.get("O365_WRITE_ENABLED", "").lower() == "true"


def require_write() -> str | None:
    """Return an error message if writes are disabled, else None."""
    if not is_write_enabled():
        return "Error: Write operations are disabled. Set O365_WRITE_ENABLED=true to enable."
    return None
