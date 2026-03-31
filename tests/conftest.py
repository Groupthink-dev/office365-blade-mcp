"""Shared fixtures for Office 365 Blade MCP tests."""

from __future__ import annotations

from unittest.mock import MagicMock, patch

import pytest


@pytest.fixture
def mock_graph_http():
    """Patch httpx.Client to return a mock."""
    with patch("office365_blade_mcp.client.httpx.Client") as mock_cls:
        mock_instance = MagicMock()
        mock_cls.return_value = mock_instance
        yield mock_instance


@pytest.fixture
def mock_acquire_token():
    """Patch acquire_token to return a fake token."""
    with patch("office365_blade_mcp.client.acquire_token") as mock_fn:
        mock_fn.return_value = "fake-access-token"
        yield mock_fn


@pytest.fixture
def client(mock_graph_http, mock_acquire_token):
    """Create a GraphClient with mocked HTTP and auth."""
    from office365_blade_mcp.client import GraphClient

    return GraphClient()


@pytest.fixture
def mock_client():
    """Patch _get_client in server to return a mock GraphClient."""
    with patch("office365_blade_mcp.server._get_client") as mock_get:
        mock_gc = MagicMock()
        mock_get.return_value = mock_gc
        yield mock_gc


# ---------------------------------------------------------------------------
# Sample data
# ---------------------------------------------------------------------------


@pytest.fixture
def sample_folders() -> list[dict]:
    return [
        {"id": "folder-inbox", "displayName": "Inbox", "totalItemCount": 1234, "unreadItemCount": 56},
        {"id": "folder-sent", "displayName": "Sent Items", "totalItemCount": 890, "unreadItemCount": 0},
        {"id": "folder-drafts", "displayName": "Drafts", "totalItemCount": 3, "unreadItemCount": 3},
    ]


@pytest.fixture
def sample_emails() -> list[dict]:
    return [
        {
            "id": "msg-001",
            "conversationId": "conv-001",
            "receivedDateTime": "2026-03-07T10:30:00Z",
            "subject": "Meeting notes",
            "from": {"emailAddress": {"name": "Alice", "address": "alice@example.com"}},
            "toRecipients": [{"emailAddress": {"name": "Bob", "address": "bob@example.com"}}],
            "isRead": True,
            "flag": {"flagStatus": "notFlagged"},
            "importance": "normal",
            "hasAttachments": False,
            "bodyPreview": "Here are the meeting notes from today...",
        },
        {
            "id": "msg-002",
            "conversationId": "conv-002",
            "receivedDateTime": "2026-03-06T14:00:00Z",
            "subject": "Re: Project update",
            "from": {"emailAddress": {"name": "Bob", "address": "bob@example.com"}},
            "toRecipients": [{"emailAddress": {"name": "Alice", "address": "alice@example.com"}}],
            "isRead": True,
            "flag": {"flagStatus": "flagged"},
            "importance": "high",
            "hasAttachments": True,
            "bodyPreview": "The project is on track...",
        },
    ]


@pytest.fixture
def sample_email_full() -> dict:
    return {
        "id": "msg-001",
        "conversationId": "conv-001",
        "receivedDateTime": "2026-03-07T10:30:00Z",
        "subject": "Meeting notes",
        "from": {"emailAddress": {"name": "Alice", "address": "alice@example.com"}},
        "toRecipients": [{"emailAddress": {"name": "Bob", "address": "bob@example.com"}}],
        "ccRecipients": [{"emailAddress": {"address": "charlie@example.com"}}],
        "isRead": True,
        "flag": {"flagStatus": "notFlagged"},
        "hasAttachments": False,
        "body": {"contentType": "Text", "content": "Here are the meeting notes from today.\n\nBest regards,\nAlice"},
        "bodyPreview": "Here are the meeting notes from today...",
    }


@pytest.fixture
def sample_calendars() -> list[dict]:
    return [
        {"id": "cal-default", "name": "Calendar", "isDefaultCalendar": True, "canEdit": True},
        {"id": "cal-work", "name": "Work Calendar", "isDefaultCalendar": False, "canEdit": True},
    ]


@pytest.fixture
def sample_events() -> list[dict]:
    return [
        {
            "id": "event-001",
            "subject": "Team standup",
            "start": {"dateTime": "2026-03-07T10:00:00.0000000", "timeZone": "UTC"},
            "end": {"dateTime": "2026-03-07T10:30:00.0000000", "timeZone": "UTC"},
            "location": {"displayName": "Room A"},
            "isAllDay": False,
            "isCancelled": False,
            "responseStatus": {"response": "accepted"},
            "showAs": "busy",
            "isOnlineMeeting": False,
        },
        {
            "id": "event-002",
            "subject": "1:1 with Alice",
            "start": {"dateTime": "2026-03-07T14:00:00.0000000", "timeZone": "UTC"},
            "end": {"dateTime": "2026-03-07T15:00:00.0000000", "timeZone": "UTC"},
            "location": {"displayName": ""},
            "isAllDay": False,
            "isCancelled": False,
            "responseStatus": {"response": "tentativelyAccepted"},
            "showAs": "tentative",
            "isOnlineMeeting": True,
            "onlineMeetingUrl": "https://teams.microsoft.com/meet/123",
        },
    ]


@pytest.fixture
def sample_task_lists() -> list[dict]:
    return [
        {"id": "tl-default", "displayName": "Tasks", "isOwner": True, "isShared": False},
        {"id": "tl-shopping", "displayName": "Shopping", "isOwner": True, "isShared": True},
    ]


@pytest.fixture
def sample_tasks() -> list[dict]:
    return [
        {
            "id": "task-001",
            "title": "Buy groceries",
            "status": "notStarted",
            "importance": "high",
            "createdDateTime": "2026-03-06T10:00:00Z",
            "dueDateTime": {"dateTime": "2026-03-08T00:00:00", "timeZone": "UTC"},
            "completedDateTime": None,
            "body": {"contentType": "text", "content": "Milk, eggs, bread"},
        },
        {
            "id": "task-002",
            "title": "Send report",
            "status": "completed",
            "importance": "normal",
            "createdDateTime": "2026-03-05T09:00:00Z",
            "dueDateTime": {"dateTime": "2026-03-06T00:00:00", "timeZone": "UTC"},
            "completedDateTime": {"dateTime": "2026-03-06T15:00:00", "timeZone": "UTC"},
            "body": {"contentType": "text", "content": ""},
        },
    ]
