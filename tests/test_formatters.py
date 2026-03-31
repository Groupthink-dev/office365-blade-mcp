"""Tests for token-efficient formatters."""

from __future__ import annotations

from office365_blade_mcp.formatters import (
    format_calendar_list,
    format_email_body,
    format_email_list,
    format_email_snippets,
    format_email_thread,
    format_event_detail,
    format_event_list,
    format_folder_list,
    format_freebusy,
    format_task_list_items,
    format_task_lists,
)

# ===========================================================================
# EMAIL FORMATTERS
# ===========================================================================


class TestFormatFolderList:
    def test_formats_folders(self, sample_folders):
        result = format_folder_list(sample_folders)
        assert "Inbox (1234/56)" in result
        assert "Sent Items (890/0)" in result
        assert "folder-inbox" in result

    def test_empty(self):
        assert "No mail folders" in format_folder_list([])


class TestFormatEmailList:
    def test_formats_emails(self, sample_emails):
        result = format_email_list(sample_emails)
        assert "Alice" in result
        assert "Meeting notes" in result
        assert "2026-03-07" in result

    def test_shows_flags(self, sample_emails):
        result = format_email_list(sample_emails)
        assert "flagged" in result
        assert "attach" in result

    def test_truncation_annotation(self, sample_emails):
        result = format_email_list(sample_emails, total=100, limit=2)
        assert "98 more" in result

    def test_empty(self):
        assert "No emails found" in format_email_list([])


class TestFormatEmailBody:
    def test_formats_full_email(self, sample_email_full):
        result = format_email_body(sample_email_full)
        assert "From: Alice <alice@example.com>" in result
        assert "Subject: Meeting notes" in result
        assert "meeting notes from today" in result

    def test_strips_html(self):
        msg = {
            "from": {"emailAddress": {"name": "Test", "address": "test@example.com"}},
            "subject": "HTML email",
            "body": {"contentType": "HTML", "content": "<html><body><p>Hello <b>world</b></p></body></html>"},
        }
        result = format_email_body(msg)
        assert "Hello world" in result
        assert "<html>" not in result


class TestFormatEmailSnippets:
    def test_formats_previews(self, sample_emails):
        result = format_email_snippets(sample_emails)
        assert "Meeting notes" in result
        assert "meeting notes from today" in result.lower()


class TestFormatEmailThread:
    def test_formats_thread(self, sample_emails):
        emails = []
        for e in sample_emails:
            ec = dict(e)
            ec["body"] = {"contentType": "Text", "content": f"Body of {e['subject']}"}
            emails.append(ec)
        result = format_email_thread(emails)
        assert "[1/2]" in result
        assert "[2/2]" in result
        assert "Meeting notes" in result


# ===========================================================================
# CALENDAR FORMATTERS
# ===========================================================================


class TestFormatCalendarList:
    def test_formats_calendars(self, sample_calendars):
        result = format_calendar_list(sample_calendars)
        assert "Calendar" in result
        assert "default" in result
        assert "editable" in result

    def test_empty(self):
        assert "No calendars" in format_calendar_list([])


class TestFormatEventList:
    def test_formats_events(self, sample_events):
        result = format_event_list(sample_events)
        assert "Team standup" in result
        assert "Room A" in result
        assert "accepted" in result
        assert "online" in result

    def test_empty(self):
        assert "No events" in format_event_list([])


class TestFormatEventDetail:
    def test_formats_detail(self, sample_events):
        result = format_event_detail(sample_events[0])
        assert "Team standup" in result
        assert "Room A" in result


class TestFormatFreebusy:
    def test_formats_schedule(self):
        schedules = [
            {
                "scheduleId": "alice@example.com",
                "availabilityView": "002000",
                "scheduleItems": [
                    {
                        "start": {"dateTime": "2026-03-07T10:00:00"},
                        "end": {"dateTime": "2026-03-07T10:30:00"},
                        "status": "busy",
                        "subject": "Meeting",
                    }
                ],
            }
        ]
        result = format_freebusy(schedules)
        assert "alice@example.com" in result
        assert "busy" in result
        assert "Meeting" in result


# ===========================================================================
# TASK FORMATTERS
# ===========================================================================


class TestFormatTaskLists:
    def test_formats_lists(self, sample_task_lists):
        result = format_task_lists(sample_task_lists)
        assert "Tasks" in result
        assert "owner" in result
        assert "Shopping" in result

    def test_empty(self):
        assert "No task lists" in format_task_lists([])


class TestFormatTaskListItems:
    def test_formats_tasks(self, sample_tasks):
        result = format_task_list_items(sample_tasks)
        assert "[ ]" in result
        assert "[x]" in result
        assert "Buy groceries" in result
        assert "high" in result
        assert "due:" in result

    def test_empty(self):
        assert "No tasks" in format_task_list_items([])
