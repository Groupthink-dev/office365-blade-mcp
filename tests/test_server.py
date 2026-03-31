"""Tests for MCP server tools."""

from __future__ import annotations

from unittest.mock import patch

from office365_blade_mcp.client import GraphError, NotFoundError

# ===========================================================================
# META TOOLS
# ===========================================================================


class TestO365Info:
    async def test_success(self, mock_client):
        from office365_blade_mcp.server import o365_info

        mock_client.get_user_info.return_value = {
            "displayName": "Test User",
            "mail": "test@example.com",
            "userPrincipalName": "test@example.com",
        }
        result = await o365_info()
        assert "Test User" in result
        assert "test@example.com" in result

    async def test_error(self, mock_client):
        from office365_blade_mcp.server import o365_info

        mock_client.get_user_info.side_effect = GraphError("Connection failed")
        result = await o365_info()
        assert "Error:" in result


# ===========================================================================
# EMAIL READ TOOLS
# ===========================================================================


class TestEmailFolders:
    async def test_success(self, mock_client, sample_folders):
        from office365_blade_mcp.server import email_folders

        mock_client.get_mail_folders.return_value = sample_folders
        result = await email_folders()
        assert "Inbox" in result
        assert "(1234/56)" in result


class TestEmailSearch:
    async def test_success(self, mock_client, sample_emails):
        from office365_blade_mcp.server import email_search

        mock_client.search_emails.return_value = (sample_emails, 2)
        result = await email_search(from_addr="alice@example.com")
        assert "Meeting notes" in result

    async def test_no_results(self, mock_client):
        from office365_blade_mcp.server import email_search

        mock_client.search_emails.return_value = ([], 0)
        result = await email_search(subject="nonexistent")
        assert "No emails found" in result

    async def test_error(self, mock_client):
        from office365_blade_mcp.server import email_search

        mock_client.search_emails.side_effect = GraphError("Timeout")
        result = await email_search(subject="test")
        assert "Error:" in result


class TestEmailRead:
    async def test_success(self, mock_client, sample_email_full):
        from office365_blade_mcp.server import email_read

        mock_client.get_email.return_value = sample_email_full
        result = await email_read(id="msg-001")
        assert "Alice" in result
        assert "Meeting notes" in result
        assert "meeting notes from today" in result

    async def test_not_found(self, mock_client):
        from office365_blade_mcp.server import email_read

        mock_client.get_email.side_effect = NotFoundError("Not found", 404)
        result = await email_read(id="nonexistent")
        assert "Error:" in result


class TestEmailThread:
    async def test_success(self, mock_client, sample_emails):
        from office365_blade_mcp.server import email_thread

        # Add body to sample emails for thread view
        emails_with_body = []
        for e in sample_emails:
            ec = dict(e)
            ec["body"] = {"contentType": "Text", "content": f"Body of {e['subject']}"}
            emails_with_body.append(ec)
        mock_client.get_email_thread.return_value = emails_with_body
        result = await email_thread(conversation_id="conv-001")
        assert "[1/2]" in result
        assert "Meeting notes" in result


class TestEmailSnippets:
    async def test_success(self, mock_client, sample_emails):
        from office365_blade_mcp.server import email_snippets

        mock_client.get_email_snippets.return_value = (sample_emails, 2)
        result = await email_snippets(from_addr="alice@example.com")
        assert "Meeting notes" in result
        assert "meeting notes from today" in result.lower()


# ===========================================================================
# EMAIL WRITE TOOLS
# ===========================================================================


class TestEmailSend:
    async def test_write_disabled(self, mock_client):
        from office365_blade_mcp.server import email_send

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "false"}):
            result = await email_send(to="bob@example.com", subject="Test", body="Hello")
            assert "Write operations are disabled" in result

    async def test_success(self, mock_client):
        from office365_blade_mcp.server import email_send

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "true"}):
            mock_client.send_email.return_value = "sent"
            result = await email_send(to="bob@example.com", subject="Test", body="Hello")
            assert "Sent" in result


class TestEmailReply:
    async def test_write_disabled(self, mock_client):
        from office365_blade_mcp.server import email_reply

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "false"}):
            result = await email_reply(id="msg-001", body="Thanks!")
            assert "Write operations are disabled" in result


class TestEmailDelete:
    async def test_requires_confirm(self, mock_client):
        from office365_blade_mcp.server import email_delete

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "true"}):
            result = await email_delete(ids="msg-001", confirm=False)
            assert "confirm=true" in result

    async def test_success_with_confirm(self, mock_client):
        from office365_blade_mcp.server import email_delete

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "true"}):
            mock_client.delete_email.return_value = "deleted"
            result = await email_delete(ids="msg-001", confirm=True)
            assert "Deleted" in result


# ===========================================================================
# CALENDAR READ TOOLS
# ===========================================================================


class TestCalCalendars:
    async def test_success(self, mock_client, sample_calendars):
        from office365_blade_mcp.server import cal_calendars

        mock_client.get_calendars.return_value = sample_calendars
        result = await cal_calendars()
        assert "Calendar" in result
        assert "default" in result


class TestCalEvents:
    async def test_success(self, mock_client, sample_events):
        from office365_blade_mcp.server import cal_events

        mock_client.get_events.return_value = sample_events
        result = await cal_events(start="2026-03-07", end="2026-03-07")
        assert "Team standup" in result
        assert "Room A" in result


class TestCalToday:
    async def test_success(self, mock_client, sample_events):
        from office365_blade_mcp.server import cal_today

        mock_client.get_events.return_value = sample_events
        result = await cal_today()
        assert "Team standup" in result


class TestCalFreebusy:
    async def test_success(self, mock_client):
        from office365_blade_mcp.server import cal_freebusy

        mock_client.get_freebusy.return_value = [
            {
                "scheduleId": "alice@example.com",
                "availabilityView": "0020000020",
                "scheduleItems": [
                    {
                        "start": {"dateTime": "2026-03-07T10:00:00"},
                        "end": {"dateTime": "2026-03-07T10:30:00"},
                        "status": "busy",
                        "subject": "Standup",
                    }
                ],
            }
        ]
        result = await cal_freebusy(start="2026-03-07", end="2026-03-07")
        assert "alice@example.com" in result
        assert "busy" in result


# ===========================================================================
# CALENDAR WRITE TOOLS
# ===========================================================================


class TestCalCreate:
    async def test_write_disabled(self, mock_client):
        from office365_blade_mcp.server import cal_create

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "false"}):
            result = await cal_create(subject="Test", start="2026-03-08T10:00:00", end="2026-03-08T11:00:00")
            assert "Write operations are disabled" in result


class TestCalDelete:
    async def test_requires_confirm(self, mock_client):
        from office365_blade_mcp.server import cal_delete

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "true"}):
            result = await cal_delete(id="event-001", confirm=False)
            assert "confirm=true" in result


class TestCalRespond:
    async def test_success(self, mock_client):
        from office365_blade_mcp.server import cal_respond

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "true"}):
            mock_client.respond_event.return_value = "responded: accept"
            result = await cal_respond(id="event-001", response="accept")
            assert "accept" in result


# ===========================================================================
# TASKS TOOLS
# ===========================================================================


class TestTaskLists:
    async def test_success(self, mock_client, sample_task_lists):
        from office365_blade_mcp.server import task_lists

        mock_client.get_task_lists.return_value = sample_task_lists
        result = await task_lists()
        assert "Tasks" in result
        assert "Shopping" in result


class TestTaskList:
    async def test_success(self, mock_client, sample_tasks):
        from office365_blade_mcp.server import task_list

        mock_client.get_tasks.return_value = sample_tasks
        result = await task_list(list_id="tl-default")
        assert "Buy groceries" in result
        assert "[ ]" in result
        assert "[x]" in result


class TestTaskCreate:
    async def test_write_disabled(self, mock_client):
        from office365_blade_mcp.server import task_create

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "false"}):
            result = await task_create(list_id="tl-default", title="Test task")
            assert "Write operations are disabled" in result

    async def test_success(self, mock_client):
        from office365_blade_mcp.server import task_create

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "true"}):
            mock_client.create_task.return_value = {"id": "task-new", "title": "Test task"}
            result = await task_create(list_id="tl-default", title="Test task")
            assert "Created task" in result


class TestTaskComplete:
    async def test_success(self, mock_client):
        from office365_blade_mcp.server import task_complete

        with patch.dict("os.environ", {"O365_WRITE_ENABLED": "true"}):
            mock_client.complete_task.return_value = {"id": "task-001", "status": "completed"}
            result = await task_complete(list_id="tl-default", task_id="task-001")
            assert "Completed" in result
