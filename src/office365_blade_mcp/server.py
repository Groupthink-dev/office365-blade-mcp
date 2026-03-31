"""Office 365 Blade MCP Server — email, calendar, and tasks via Microsoft Graph API.

Wraps the Microsoft Graph API as MCP tools. Token-efficient by default:
concise output, field selection, capped lists, null-field omission.
"""

from __future__ import annotations

import asyncio
import logging
import os
from datetime import UTC, datetime, timedelta
from typing import Annotated, Any

from fastmcp import FastMCP
from pydantic import Field

from office365_blade_mcp.client import GraphClient, GraphError
from office365_blade_mcp.formatters import (
    format_attachments,
    format_calendar_list,
    format_email_body,
    format_email_changes,
    format_email_list,
    format_email_snippets,
    format_email_thread,
    format_event_detail,
    format_event_list,
    format_folder_list,
    format_freebusy,
    format_planner_plans,
    format_planner_tasks,
    format_task_list_items,
    format_task_lists,
    format_user_info,
)
from office365_blade_mcp.models import DEFAULT_LIMIT, MAX_BATCH_SIZE, require_write

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Transport configuration
# ---------------------------------------------------------------------------

TRANSPORT = os.environ.get("O365_MCP_TRANSPORT", "stdio")
HTTP_HOST = os.environ.get("O365_MCP_HOST", "127.0.0.1")
HTTP_PORT = int(os.environ.get("O365_MCP_PORT", "8770"))

# ---------------------------------------------------------------------------
# FastMCP server
# ---------------------------------------------------------------------------

mcp = FastMCP("Office365Blade")

# Lazy-initialized client
_client: GraphClient | None = None


def _get_client() -> GraphClient:
    """Get or create the GraphClient singleton."""
    global _client  # noqa: PLW0603
    if _client is None:
        _client = GraphClient()
        logger.info("GraphClient initialised")
    return _client


def _error_response(e: GraphError) -> str:
    """Format a client error as a user-friendly string."""
    return f"Error: {e}"


async def _run(fn: Any, *args: Any, **kwargs: Any) -> Any:
    """Run a blocking client method in a thread to avoid blocking the event loop."""
    return await asyncio.to_thread(fn, *args, **kwargs)


# ===========================================================================
# META TOOLS
# ===========================================================================


@mcp.tool
async def o365_info() -> str:
    """Get Microsoft 365 account info: name, email, UPN.

    Use this as a health check and to confirm account connectivity.
    """
    try:
        result = await _run(_get_client().get_user_info)
        return format_user_info(result)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in o365_info")
        return f"Error: {e}"


# ===========================================================================
# EMAIL READ TOOLS
# ===========================================================================


@mcp.tool
async def email_folders() -> str:
    """List all mail folders with ID, name, total/unread counts.

    Returns Inbox, Sent Items, Drafts, Deleted Items, and custom folders.
    """
    try:
        result = await _run(_get_client().get_mail_folders)
        return format_folder_list(result)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_folders")
        return f"Error: {e}"


@mcp.tool
async def email_search(
    from_addr: Annotated[str | None, Field(description="Filter by sender email address")] = None,
    to_addr: Annotated[str | None, Field(description="Filter by recipient email address")] = None,
    subject: Annotated[str | None, Field(description="Filter by subject text (partial match)")] = None,
    body: Annotated[str | None, Field(description="Filter by body text (partial match)")] = None,
    after: Annotated[str | None, Field(description="Emails after this date (YYYY-MM-DD)")] = None,
    before: Annotated[str | None, Field(description="Emails before this date (YYYY-MM-DD)")] = None,
    folder_id: Annotated[str | None, Field(description="Folder ID (from email_folders)")] = None,
    is_read: Annotated[bool | None, Field(description="Filter by read status")] = None,
    has_attachments: Annotated[bool | None, Field(description="Filter by attachment presence")] = None,
    limit: Annotated[int, Field(description="Max results (default: 20)")] = DEFAULT_LIMIT,
) -> str:
    """Search emails with filters. Returns concise list: date, sender, subject, flags.

    At least one filter should be provided. Use ``folder_id`` from ``email_folders``.
    """
    try:
        emails, total = await _run(
            _get_client().search_emails,
            from_addr=from_addr,
            to_addr=to_addr,
            subject=subject,
            body=body,
            after=after,
            before=before,
            folder_id=folder_id,
            is_read=is_read,
            has_attachments=has_attachments,
            limit=limit,
        )
        return format_email_list(emails, total=total, limit=limit)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_search")
        return f"Error: {e}"


@mcp.tool
async def email_read(
    id: Annotated[str, Field(description="Email message ID")],
) -> str:
    """Read a full email: headers + body.

    Returns From, To, Cc, Subject, Date, flags, and body text.
    """
    try:
        msg = await _run(_get_client().get_email, id)
        return format_email_body(msg)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_read")
        return f"Error: {e}"


@mcp.tool
async def email_thread(
    conversation_id: Annotated[str, Field(description="Conversation ID (from email_search or email_read)")],
    limit: Annotated[int, Field(description="Max messages (default: 20)")] = DEFAULT_LIMIT,
) -> str:
    """Get all messages in a conversation thread, ordered chronologically.

    Thread IDs are returned by ``email_search`` and ``email_read``.
    """
    try:
        messages = await _run(_get_client().get_email_thread, conversation_id, limit=limit)
        return format_email_thread(messages)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_thread")
        return f"Error: {e}"


@mcp.tool
async def email_snippets(
    from_addr: Annotated[str | None, Field(description="Filter by sender email")] = None,
    subject: Annotated[str | None, Field(description="Filter by subject text")] = None,
    after: Annotated[str | None, Field(description="After date (YYYY-MM-DD)")] = None,
    before: Annotated[str | None, Field(description="Before date (YYYY-MM-DD)")] = None,
    folder_id: Annotated[str | None, Field(description="Folder ID")] = None,
    limit: Annotated[int, Field(description="Max results (default: 20)")] = DEFAULT_LIMIT,
) -> str:
    """Search emails returning preview snippets only (90% fewer tokens than full read).

    Best for browsing and finding specific content before reading full emails.
    """
    try:
        emails, total = await _run(
            _get_client().get_email_snippets,
            from_addr=from_addr,
            subject=subject,
            after=after,
            before=before,
            folder_id=folder_id,
            limit=limit,
        )
        return format_email_snippets(emails, total=total, limit=limit)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_snippets")
        return f"Error: {e}"


@mcp.tool
async def email_attachments(
    id: Annotated[str, Field(description="Email message ID")],
) -> str:
    """List attachments for an email: name, type, size.

    Returns metadata only — does not download attachment content.
    """
    try:
        attachments = await _run(_get_client().get_email_attachments, id)
        return format_attachments(attachments)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_attachments")
        return f"Error: {e}"


# ===========================================================================
# EMAIL STATE & CHANGES
# ===========================================================================


@mcp.tool
async def email_state() -> str:
    """Get a delta link for email change tracking.

    Returns a delta link that can be passed to ``email_changes`` to get
    incremental updates. Use this to initialise a watermark for change tracking.
    """
    try:
        delta_link = await _run(_get_client().get_email_state)
        return f"Delta link: {delta_link}"
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_state")
        return f"Error: {e}"


@mcp.tool
async def email_changes(
    delta_link: Annotated[str, Field(description="Delta link from email_state or previous email_changes")],
) -> str:
    """Get email changes since a delta link. Returns changed/removed message IDs.

    Use ``email_state`` to get the initial delta link.
    """
    try:
        changes = await _run(_get_client().get_email_changes, delta_link)
        return format_email_changes(changes)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_changes")
        return f"Error: {e}"


# ===========================================================================
# EMAIL WRITE TOOLS (gated: O365_WRITE_ENABLED=true)
# ===========================================================================


@mcp.tool
async def email_send(
    to: Annotated[str, Field(description="Recipient email(s), comma-separated")],
    subject: Annotated[str, Field(description="Email subject")],
    body: Annotated[str, Field(description="Email body (plain text)")],
    cc: Annotated[str | None, Field(description="CC recipient(s), comma-separated")] = None,
    bcc: Annotated[str | None, Field(description="BCC recipient(s), comma-separated")] = None,
) -> str:
    """Send a new email. Requires O365_WRITE_ENABLED=true."""
    if err := require_write():
        return err
    try:
        to_list = [addr.strip() for addr in to.split(",") if addr.strip()]
        cc_list = [addr.strip() for addr in cc.split(",") if addr.strip()] if cc else None
        bcc_list = [addr.strip() for addr in bcc.split(",") if addr.strip()] if bcc else None
        logger.info("email_send: to=%s, subject=%s", to_list, subject[:50])
        result = await _run(_get_client().send_email, to_list, subject, body, cc_list, bcc_list)
        return f"Sent: {result}"
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_send")
        return f"Error: {e}"


@mcp.tool
async def email_reply(
    id: Annotated[str, Field(description="Email message ID to reply to")],
    body: Annotated[str, Field(description="Reply body (plain text)")],
    reply_all: Annotated[bool, Field(description="Reply to all recipients (default: false)")] = False,
) -> str:
    """Reply to an email. Preserves threading. Requires O365_WRITE_ENABLED=true."""
    if err := require_write():
        return err
    try:
        logger.info("email_reply: id=%s, reply_all=%s", id[:20], reply_all)
        result = await _run(_get_client().reply_email, id, body, reply_all)
        return f"Reply sent: {result}"
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_reply")
        return f"Error: {e}"


@mcp.tool
async def email_forward(
    id: Annotated[str, Field(description="Email message ID to forward")],
    to: Annotated[str, Field(description="Forward to email(s), comma-separated")],
    comment: Annotated[str, Field(description="Optional comment to prepend")] = "",
) -> str:
    """Forward an email. Requires O365_WRITE_ENABLED=true."""
    if err := require_write():
        return err
    try:
        to_list = [addr.strip() for addr in to.split(",") if addr.strip()]
        logger.info("email_forward: id=%s, to=%s", id[:20], to_list)
        result = await _run(_get_client().forward_email, id, to_list, comment)
        return f"Forwarded: {result}"
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_forward")
        return f"Error: {e}"


@mcp.tool
async def email_flag(
    ids: Annotated[str, Field(description="Email ID(s), comma-separated")],
    action: Annotated[str, Field(description="Action: flag, unflag, mark_read, mark_unread")] = "flag",
) -> str:
    """Flag/unflag or mark read/unread on emails. Requires O365_WRITE_ENABLED=true.

    Actions: ``flag``, ``unflag``, ``mark_read``, ``mark_unread``.
    """
    if err := require_write():
        return err
    try:
        id_list = [eid.strip() for eid in ids.split(",") if eid.strip()]
        logger.info("email_flag: %s on %d emails", action, len(id_list))
        count = 0
        for msg_id in id_list[:MAX_BATCH_SIZE]:
            if action == "flag":
                await _run(_get_client().flag_email, msg_id, "flagged")
            elif action == "unflag":
                await _run(_get_client().flag_email, msg_id, "notFlagged")
            elif action == "mark_read":
                await _run(_get_client().mark_email_read, msg_id, True)
            elif action == "mark_unread":
                await _run(_get_client().mark_email_read, msg_id, False)
            else:
                return f"Error: Unknown action '{action}'. Valid: flag, unflag, mark_read, mark_unread"
            count += 1
        return f"{action}: {count} email(s) affected."
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_flag")
        return f"Error: {e}"


@mcp.tool
async def email_move(
    ids: Annotated[str, Field(description="Email ID(s), comma-separated")],
    folder_id: Annotated[str, Field(description="Destination folder ID (from email_folders)")],
) -> str:
    """Move emails to a folder. Requires O365_WRITE_ENABLED=true.

    Use ``email_folders`` to find folder IDs.
    """
    if err := require_write():
        return err
    try:
        id_list = [eid.strip() for eid in ids.split(",") if eid.strip()]
        logger.info("email_move: %d emails to %s", len(id_list), folder_id[:20])
        count = 0
        for msg_id in id_list[:MAX_BATCH_SIZE]:
            await _run(_get_client().move_email, msg_id, folder_id)
            count += 1
        return f"Moved {count} email(s)."
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_move")
        return f"Error: {e}"


@mcp.tool
async def email_delete(
    ids: Annotated[str, Field(description="Email ID(s), comma-separated")],
    confirm: Annotated[bool, Field(description="Must be true to confirm deletion")] = False,
) -> str:
    """Delete emails (moves to Deleted Items). Requires O365_WRITE_ENABLED=true.

    Set ``confirm=true`` to confirm. This is a safety gate — deletion moves
    messages to Deleted Items (recoverable), not permanent destroy.
    """
    if err := require_write():
        return err
    if not confirm:
        return "Error: Set confirm=true to confirm deletion. Messages will be moved to Deleted Items."
    try:
        id_list = [eid.strip() for eid in ids.split(",") if eid.strip()]
        logger.info("email_delete: %d emails", len(id_list))
        count = 0
        for msg_id in id_list[:MAX_BATCH_SIZE]:
            await _run(_get_client().delete_email, msg_id)
            count += 1
        return f"Deleted {count} email(s) (moved to Deleted Items)."
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_delete")
        return f"Error: {e}"


@mcp.tool
async def email_bulk(
    ids: Annotated[str, Field(description="Email ID(s), comma-separated (max 50)")],
    action: Annotated[str, Field(description="Action: mark_read, mark_unread, flag, unflag, move, delete")],
    target_folder: Annotated[str | None, Field(description="Target folder ID (required for 'move')")] = None,
) -> str:
    """Bulk action on emails. Capped at 50. Requires O365_WRITE_ENABLED=true."""
    if err := require_write():
        return err
    try:
        id_list = [eid.strip() for eid in ids.split(",") if eid.strip()]
        if len(id_list) > MAX_BATCH_SIZE:
            return f"Error: Maximum {MAX_BATCH_SIZE} emails per bulk operation. Got {len(id_list)}."
        logger.info("email_bulk: %s on %d emails", action, len(id_list))
        count = await _run(_get_client().bulk_email_action, id_list, action, target_folder)
        return f"Bulk {action}: {count} email(s) affected."
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in email_bulk")
        return f"Error: {e}"


# ===========================================================================
# CALENDAR READ TOOLS
# ===========================================================================


@mcp.tool
async def cal_calendars() -> str:
    """List all calendars with ID, name, and permissions."""
    try:
        result = await _run(_get_client().get_calendars)
        return format_calendar_list(result)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in cal_calendars")
        return f"Error: {e}"


@mcp.tool
async def cal_events(
    start: Annotated[str, Field(description="Start date (YYYY-MM-DD or ISO datetime)")],
    end: Annotated[str, Field(description="End date (YYYY-MM-DD or ISO datetime)")],
    calendar_id: Annotated[str | None, Field(description="Calendar ID (from cal_calendars)")] = None,
    limit: Annotated[int, Field(description="Max results (default: 20)")] = DEFAULT_LIMIT,
) -> str:
    """Get calendar events in a date range.

    Returns time, subject, location, response status.
    """
    try:
        events = await _run(_get_client().get_events, start, end, calendar_id, limit)
        return format_event_list(events)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in cal_events")
        return f"Error: {e}"


@mcp.tool
async def cal_event(
    id: Annotated[str, Field(description="Event ID")],
) -> str:
    """Get full detail for a single event: attendees, body, recurrence."""
    try:
        event = await _run(_get_client().get_event, id)
        return format_event_detail(event)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in cal_event")
        return f"Error: {e}"


@mcp.tool
async def cal_search(
    query: Annotated[str, Field(description="Search text (matches subject)")],
    start: Annotated[str | None, Field(description="Start date filter (YYYY-MM-DD)")] = None,
    end: Annotated[str | None, Field(description="End date filter (YYYY-MM-DD)")] = None,
    limit: Annotated[int, Field(description="Max results (default: 20)")] = DEFAULT_LIMIT,
) -> str:
    """Search events by subject text."""
    try:
        events = await _run(_get_client().search_events, query, start, end, limit)
        return format_event_list(events)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in cal_search")
        return f"Error: {e}"


@mcp.tool
async def cal_today() -> str:
    """Get all events for today. Convenience shortcut."""
    try:
        today = datetime.now(UTC).strftime("%Y-%m-%d")
        events = await _run(_get_client().get_events, today, today, None, 50)
        return format_event_list(events)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in cal_today")
        return f"Error: {e}"


@mcp.tool
async def cal_week() -> str:
    """Get all events for this week (Mon-Sun). Convenience shortcut."""
    try:
        now = datetime.now(UTC)
        monday = now - timedelta(days=now.weekday())
        sunday = monday + timedelta(days=6)
        events = await _run(
            _get_client().get_events,
            monday.strftime("%Y-%m-%d"),
            sunday.strftime("%Y-%m-%d"),
            None,
            100,
        )
        return format_event_list(events)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in cal_week")
        return f"Error: {e}"


@mcp.tool
async def cal_freebusy(
    start: Annotated[str, Field(description="Start date (YYYY-MM-DD)")],
    end: Annotated[str, Field(description="End date (YYYY-MM-DD)")],
    schedules: Annotated[str | None, Field(description="Email(s) to check, comma-separated (default: self)")] = None,
) -> str:
    """Get free/busy availability for users.

    Returns availability view and busy time slots.
    """
    try:
        schedule_list = [s.strip() for s in schedules.split(",") if s.strip()] if schedules else None
        result = await _run(_get_client().get_freebusy, start, end, schedule_list)
        return format_freebusy(result)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in cal_freebusy")
        return f"Error: {e}"


@mcp.tool
async def cal_batch(
    start: Annotated[str, Field(description="Start date (YYYY-MM-DD)")],
    end: Annotated[str, Field(description="End date (YYYY-MM-DD)")],
    calendar_ids: Annotated[str, Field(description="Calendar IDs, comma-separated")],
    limit: Annotated[int, Field(description="Max events per calendar (default: 20)")] = DEFAULT_LIMIT,
) -> str:
    """Get events from multiple calendars in one call. More efficient than repeated cal_events."""
    try:
        ids = [cid.strip() for cid in calendar_ids.split(",") if cid.strip()]
        all_parts: list[str] = []
        for cal_id in ids:
            events = await _run(_get_client().get_events, start, end, cal_id, limit)
            all_parts.append(f"--- Calendar {cal_id[:20]}... ---")
            all_parts.append(format_event_list(events))
            all_parts.append("")
        return "\n".join(all_parts).rstrip()
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in cal_batch")
        return f"Error: {e}"


# ===========================================================================
# CALENDAR WRITE TOOLS (gated)
# ===========================================================================


@mcp.tool
async def cal_respond(
    id: Annotated[str, Field(description="Event ID")],
    response: Annotated[str, Field(description="Response: accept, decline, tentativelyAccept")],
    comment: Annotated[str, Field(description="Optional comment")] = "",
) -> str:
    """Respond to an event invitation. Requires O365_WRITE_ENABLED=true."""
    if err := require_write():
        return err
    try:
        result = await _run(_get_client().respond_event, id, response, comment)
        return result
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in cal_respond")
        return f"Error: {e}"


@mcp.tool
async def cal_create(
    subject: Annotated[str, Field(description="Event subject")],
    start: Annotated[str, Field(description="Start datetime (ISO 8601, e.g. 2026-03-08T10:00:00)")],
    end: Annotated[str, Field(description="End datetime (ISO 8601)")],
    body: Annotated[str | None, Field(description="Event body/description")] = None,
    location: Annotated[str | None, Field(description="Location name")] = None,
    attendees: Annotated[str | None, Field(description="Attendee emails, comma-separated")] = None,
    is_all_day: Annotated[bool, Field(description="All-day event")] = False,
    calendar_id: Annotated[str | None, Field(description="Calendar ID (default calendar if omitted)")] = None,
) -> str:
    """Create a calendar event. Requires O365_WRITE_ENABLED=true."""
    if err := require_write():
        return err
    try:
        attendee_list = [a.strip() for a in attendees.split(",") if a.strip()] if attendees else None
        logger.info("cal_create: %s at %s", subject, start)
        event = await _run(
            _get_client().create_event, subject, start, end, body, location, attendee_list, is_all_day, calendar_id
        )
        eid = event.get("id", "?")
        return f"Created event: {subject} (id={eid[:20]}...)"
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in cal_create")
        return f"Error: {e}"


@mcp.tool
async def cal_update(
    id: Annotated[str, Field(description="Event ID")],
    subject: Annotated[str | None, Field(description="New subject")] = None,
    start: Annotated[str | None, Field(description="New start datetime")] = None,
    end: Annotated[str | None, Field(description="New end datetime")] = None,
    location: Annotated[str | None, Field(description="New location")] = None,
    body: Annotated[str | None, Field(description="New body text")] = None,
) -> str:
    """Update a calendar event. Requires O365_WRITE_ENABLED=true."""
    if err := require_write():
        return err
    try:
        updates: dict[str, Any] = {}
        if subject:
            updates["subject"] = subject
        if start:
            updates["start"] = {"dateTime": start, "timeZone": "UTC"}
        if end:
            updates["end"] = {"dateTime": end, "timeZone": "UTC"}
        if location:
            updates["location"] = {"displayName": location}
        if body:
            updates["body"] = {"contentType": "Text", "content": body}
        if not updates:
            return "Error: No fields to update."
        logger.info("cal_update: id=%s", id[:20])
        await _run(_get_client().update_event, id, updates)
        return f"Updated event: id={id[:20]}..."
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in cal_update")
        return f"Error: {e}"


@mcp.tool
async def cal_delete(
    id: Annotated[str, Field(description="Event ID")],
    confirm: Annotated[bool, Field(description="Must be true to confirm deletion")] = False,
) -> str:
    """Delete a calendar event. Requires O365_WRITE_ENABLED=true."""
    if err := require_write():
        return err
    if not confirm:
        return "Error: Set confirm=true to confirm event deletion."
    try:
        logger.info("cal_delete: id=%s", id[:20])
        await _run(_get_client().delete_event, id)
        return "Event deleted."
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in cal_delete")
        return f"Error: {e}"


# ===========================================================================
# TASKS (To Do) READ TOOLS
# ===========================================================================


@mcp.tool
async def task_lists() -> str:
    """List all To Do task lists with ID, name, and ownership."""
    try:
        result = await _run(_get_client().get_task_lists)
        return format_task_lists(result)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in task_lists")
        return f"Error: {e}"


@mcp.tool
async def task_list(
    list_id: Annotated[str, Field(description="Task list ID (from task_lists)")],
    status: Annotated[str | None, Field(description="Filter: notStarted, inProgress, completed")] = None,
    limit: Annotated[int, Field(description="Max results (default: 20)")] = DEFAULT_LIMIT,
) -> str:
    """Get tasks from a task list."""
    try:
        tasks = await _run(_get_client().get_tasks, list_id, status, limit)
        return format_task_list_items(tasks)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in task_list")
        return f"Error: {e}"


@mcp.tool
async def task_search(
    list_id: Annotated[str, Field(description="Task list ID")],
    query: Annotated[str, Field(description="Search text (matches title)")],
    limit: Annotated[int, Field(description="Max results (default: 20)")] = DEFAULT_LIMIT,
) -> str:
    """Search tasks by title text."""
    try:
        tasks = await _run(_get_client().search_tasks, list_id, query, limit)
        return format_task_list_items(tasks)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in task_search")
        return f"Error: {e}"


@mcp.tool
async def task_today(
    list_id: Annotated[str, Field(description="Task list ID")],
) -> str:
    """Get tasks due today from a list. Convenience shortcut."""
    try:
        tasks = await _run(_get_client().get_tasks, list_id, None, 100)
        today = datetime.now(UTC).strftime("%Y-%m-%d")
        due_today = [
            t for t in tasks
            if t.get("dueDateTime", {}).get("dateTime", "")[:10] == today
        ]
        return format_task_list_items(due_today)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in task_today")
        return f"Error: {e}"


@mcp.tool
async def task_inbox() -> str:
    """Get tasks from the default task list (inbox). Convenience shortcut."""
    try:
        lists = await _run(_get_client().get_task_lists)
        # Find the default list (usually named "Tasks" and is owner)
        default_list = None
        for tl in lists:
            if tl.get("isOwner") and not tl.get("isShared"):
                default_list = tl
                break
        if not default_list and lists:
            default_list = lists[0]
        if not default_list:
            return "No task lists found."
        tasks = await _run(_get_client().get_tasks, default_list["id"], "notStarted", 50)
        return format_task_list_items(tasks)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in task_inbox")
        return f"Error: {e}"


# ===========================================================================
# TASKS WRITE TOOLS (gated)
# ===========================================================================


@mcp.tool
async def task_create(
    list_id: Annotated[str, Field(description="Task list ID")],
    title: Annotated[str, Field(description="Task title")],
    body: Annotated[str | None, Field(description="Task body/notes")] = None,
    due_date: Annotated[str | None, Field(description="Due date (YYYY-MM-DD)")] = None,
    importance: Annotated[str, Field(description="Importance: low, normal, high")] = "normal",
) -> str:
    """Create a task in a To Do list. Requires O365_WRITE_ENABLED=true."""
    if err := require_write():
        return err
    try:
        logger.info("task_create: %s in %s", title, list_id[:20])
        task = await _run(_get_client().create_task, list_id, title, body, due_date, importance)
        tid = task.get("id", "?")
        return f"Created task: {title} (id={tid[:20]}...)"
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in task_create")
        return f"Error: {e}"


@mcp.tool
async def task_update(
    list_id: Annotated[str, Field(description="Task list ID")],
    task_id: Annotated[str, Field(description="Task ID")],
    title: Annotated[str | None, Field(description="New title")] = None,
    body: Annotated[str | None, Field(description="New body text")] = None,
    due_date: Annotated[str | None, Field(description="New due date (YYYY-MM-DD)")] = None,
    importance: Annotated[str | None, Field(description="New importance: low, normal, high")] = None,
) -> str:
    """Update a task. Requires O365_WRITE_ENABLED=true."""
    if err := require_write():
        return err
    try:
        updates: dict[str, Any] = {}
        if title:
            updates["title"] = title
        if body:
            updates["body"] = {"contentType": "text", "content": body}
        if due_date:
            updates["dueDateTime"] = {"dateTime": f"{due_date}T00:00:00", "timeZone": "UTC"}
        if importance:
            updates["importance"] = importance
        if not updates:
            return "Error: No fields to update."
        logger.info("task_update: %s", task_id[:20])
        await _run(_get_client().update_task, list_id, task_id, updates)
        return f"Updated task: id={task_id[:20]}..."
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in task_update")
        return f"Error: {e}"


@mcp.tool
async def task_complete(
    list_id: Annotated[str, Field(description="Task list ID")],
    task_id: Annotated[str, Field(description="Task ID")],
) -> str:
    """Mark a task as completed. Requires O365_WRITE_ENABLED=true."""
    if err := require_write():
        return err
    try:
        logger.info("task_complete: %s", task_id[:20])
        await _run(_get_client().complete_task, list_id, task_id)
        return f"Completed task: id={task_id[:20]}..."
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in task_complete")
        return f"Error: {e}"


# ===========================================================================
# PLANNER TOOLS
# ===========================================================================


@mcp.tool
async def planner_plans() -> str:
    """List Planner plans the user has access to."""
    try:
        result = await _run(_get_client().get_planner_plans)
        return format_planner_plans(result)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in planner_plans")
        return f"Error: {e}"


@mcp.tool
async def planner_tasks(
    plan_id: Annotated[str, Field(description="Planner plan ID")],
    limit: Annotated[int, Field(description="Max results (default: 20)")] = DEFAULT_LIMIT,
) -> str:
    """Get tasks from a Planner plan."""
    try:
        tasks = await _run(_get_client().get_planner_tasks, plan_id, limit)
        return format_planner_tasks(tasks)
    except GraphError as e:
        return _error_response(e)
    except Exception as e:
        logger.exception("Unexpected error in planner_tasks")
        return f"Error: {e}"


# ===========================================================================
# Entry point
# ===========================================================================


def main() -> None:
    """Main entry point for the Office 365 Blade MCP server."""
    if TRANSPORT == "http":
        from starlette.middleware import Middleware

        from office365_blade_mcp.auth import BearerAuthMiddleware, get_bearer_token

        bearer = get_bearer_token()
        logger.info("Starting HTTP transport on %s:%s", HTTP_HOST, HTTP_PORT)
        if bearer:
            logger.info("Bearer token auth enabled")
        else:
            logger.info("Bearer token auth disabled (no O365_MCP_API_TOKEN)")
        mcp.run(
            transport="http",
            host=HTTP_HOST,
            port=HTTP_PORT,
            middleware=[Middleware(BearerAuthMiddleware)],
        )
    else:
        mcp.run()
