"""Token-efficient output formatters for Microsoft Graph API data.

Design principles:
- Concise by default (one line per item, pipe-delimited)
- Null fields omitted
- Lists capped and annotated with total count
- HTML bodies stripped to plain text
"""

from __future__ import annotations

import re
from typing import Any

from office365_blade_mcp.models import DEFAULT_LIMIT, MAX_BODY_CHARS

# ===========================================================================
# EMAIL FORMATTERS
# ===========================================================================


def format_folder_list(folders: list[dict[str, Any]]) -> str:
    """Format mail folder list: name (total/unread) id=xxx.

    Example::

        Inbox (1234/56) id=AAMk...
        Sent Items (890/0) id=AAMk...
    """
    if not folders:
        return "No mail folders found."

    lines: list[str] = []
    for f in folders:
        name = f.get("displayName", "?")
        total = f.get("totalItemCount", 0)
        unread = f.get("unreadItemCount", 0)
        fid = f.get("id", "?")
        lines.append(f"{name} ({total}/{unread}) id={fid}")
    return "\n".join(lines)


def format_email_list(emails: list[dict[str, Any]], total: int | None = None, limit: int = DEFAULT_LIMIT) -> str:
    """Format email list: date | from | subject | flags.

    Example::

        2026-03-07 10:30 | alice@example.com | Meeting notes | read
        2026-03-06 14:00 | bob@example.com | Re: Project update | read flagged
        ... 48 more (use limit= to see more)
    """
    if not emails:
        return "No emails found."

    actual_total = total if total is not None else len(emails)
    shown = emails[:limit]
    lines: list[str] = []

    for msg in shown:
        parts: list[str] = []

        # Date
        received = msg.get("receivedDateTime", "")
        parts.append(_compact_datetime(received))

        # Sender
        from_info = msg.get("from", {})
        email_addr = from_info.get("emailAddress", {})
        sender = email_addr.get("name") or email_addr.get("address") or "?"
        parts.append(sender)

        # Subject
        subject = msg.get("subject", "(no subject)")
        if len(subject) > 60:
            subject = subject[:57] + "..."
        parts.append(subject)

        # Flags
        flags: list[str] = []
        if msg.get("isRead"):
            flags.append("read")
        flag_status = msg.get("flag", {}).get("flagStatus", "")
        if flag_status == "flagged":
            flags.append("flagged")
        if msg.get("hasAttachments"):
            flags.append("attach")
        importance = msg.get("importance", "")
        if importance == "high":
            flags.append("high")
        if flags:
            parts.append(" ".join(flags))

        # ID
        mid = msg.get("id", "")
        if mid:
            parts.append(f"id={mid[:20]}...")

        lines.append(" | ".join(parts))

    if actual_total > len(shown):
        lines.append(f"... {actual_total - len(shown)} more (use limit= to see more)")

    return "\n".join(lines)


def format_email_body(msg: dict[str, Any]) -> str:
    """Format a full email for reading: headers + body.

    Example::

        From: Alice <alice@example.com>
        To: Bob <bob@example.com>
        Subject: Meeting notes
        Date: 2026-03-07T10:30:00Z
        ID: AAMk...

        Body of the email here...
    """
    lines: list[str] = []

    # Headers
    from_info = msg.get("from", {}).get("emailAddress", {})
    from_str = _format_email_address(from_info)
    if from_str:
        lines.append(f"From: {from_str}")

    to_list = msg.get("toRecipients", [])
    if to_list:
        lines.append(f"To: {', '.join(_format_email_address(r.get('emailAddress', {})) for r in to_list)}")

    cc_list = msg.get("ccRecipients", [])
    if cc_list:
        lines.append(f"Cc: {', '.join(_format_email_address(r.get('emailAddress', {})) for r in cc_list)}")

    if subject := msg.get("subject"):
        lines.append(f"Subject: {subject}")
    if date := msg.get("receivedDateTime"):
        lines.append(f"Date: {date}")
    if mid := msg.get("id"):
        lines.append(f"ID: {mid}")
    if cid := msg.get("conversationId"):
        lines.append(f"Thread: {cid}")

    # Flags
    flags: list[str] = []
    if msg.get("isRead"):
        flags.append("read")
    flag_status = msg.get("flag", {}).get("flagStatus", "")
    if flag_status == "flagged":
        flags.append("flagged")
    if msg.get("hasAttachments"):
        flags.append("attachments")
    if flags:
        lines.append(f"Flags: {' '.join(flags)}")

    lines.append("")  # Blank line before body

    # Body
    body_data = msg.get("body", {})
    content = body_data.get("content", "")
    content_type = body_data.get("contentType", "text")
    if content_type.lower() == "html":
        content = _strip_html(content)
    if content:
        lines.append(_truncate_body(content))
    else:
        preview = msg.get("bodyPreview", "")
        lines.append(preview if preview else "(no body)")

    return "\n".join(lines)


def format_email_snippets(
    emails: list[dict[str, Any]], total: int | None = None, limit: int = DEFAULT_LIMIT
) -> str:
    """Format email preview snippets (token-efficient).

    Example::

        2026-03-07 | alice@example.com | Meeting notes
          Here are the meeting notes from today...
    """
    if not emails:
        return "No emails found."

    actual_total = total if total is not None else len(emails)
    shown = emails[:limit]
    parts: list[str] = []

    for msg in shown:
        header: list[str] = []
        header.append(_compact_datetime(msg.get("receivedDateTime", "")))

        from_info = msg.get("from", {}).get("emailAddress", {})
        header.append(from_info.get("name") or from_info.get("address") or "?")
        header.append(msg.get("subject", "(no subject)"))
        header.append(f"id={msg.get('id', '?')[:20]}...")

        parts.append(" | ".join(header))
        preview = msg.get("bodyPreview", "")
        if preview:
            parts.append(f"  {preview[:200]}")
        parts.append("")

    if actual_total > len(shown):
        parts.append(f"... {actual_total - len(shown)} more")

    return "\n".join(parts).rstrip()


def format_email_thread(messages: list[dict[str, Any]]) -> str:
    """Format a conversation thread chronologically.

    Example::

        [1/3] 2026-03-05 09:00 | alice@example.com | Meeting tomorrow
        Let's meet at 2pm...
    """
    if not messages:
        return "Empty thread."

    parts: list[str] = []
    total = len(messages)

    for i, msg in enumerate(messages, 1):
        header: list[str] = [f"[{i}/{total}]"]
        header.append(_compact_datetime(msg.get("receivedDateTime", "")))

        from_info = msg.get("from", {}).get("emailAddress", {})
        header.append(from_info.get("name") or from_info.get("address") or "?")
        header.append(msg.get("subject", "(no subject)"))

        parts.append(" | ".join(header))

        body_data = msg.get("body", {})
        content = body_data.get("content", "")
        content_type = body_data.get("contentType", "text")
        if content_type.lower() == "html":
            content = _strip_html(content)
        if content:
            truncated = content[:2000]
            if len(content) > 2000:
                truncated += "\n... (truncated)"
            parts.append(truncated)
        parts.append("")

    return "\n".join(parts).rstrip()


def format_email_changes(changes: dict[str, Any]) -> str:
    """Format email delta changes."""
    items = changes.get("changes", [])
    new_delta = changes.get("new_delta_link", "")
    count = changes.get("count", 0)

    lines = [f"Changes: {count}"]
    if new_delta:
        lines.append(f"New delta link: {new_delta[:80]}...")

    for item in items[:20]:
        reason = item.get("@removed", {}).get("reason", "")
        if reason:
            lines.append(f"  Removed ({reason}): id={item.get('id', '?')[:20]}...")
        else:
            lines.append(f"  Changed: id={item.get('id', '?')[:20]}...")

    if count > 20:
        lines.append(f"  ... {count - 20} more")

    return "\n".join(lines)


def format_attachments(attachments: list[dict[str, Any]]) -> str:
    """Format attachment list."""
    if not attachments:
        return "No attachments."

    lines: list[str] = []
    for att in attachments:
        name = att.get("name", "?")
        content_type = att.get("contentType", "?")
        size = att.get("size", 0)
        aid = att.get("id", "?")
        lines.append(f"{name} | {content_type} | {_human_size(size)} | id={aid[:20]}...")
    return "\n".join(lines)


# ===========================================================================
# CALENDAR FORMATTERS
# ===========================================================================


def format_calendar_list(calendars: list[dict[str, Any]]) -> str:
    """Format calendar list.

    Example::

        Calendar (default, editable) id=AAMk...
        Work Calendar (editable) id=AAMk...
    """
    if not calendars:
        return "No calendars found."

    lines: list[str] = []
    for cal in calendars:
        name = cal.get("name", "?")
        tags: list[str] = []
        if cal.get("isDefaultCalendar"):
            tags.append("default")
        if cal.get("canEdit"):
            tags.append("editable")
        tag_str = f" ({', '.join(tags)})" if tags else ""
        cid = cal.get("id", "?")
        lines.append(f"{name}{tag_str} id={cid[:20]}...")
    return "\n".join(lines)


def format_event_list(events: list[dict[str, Any]]) -> str:
    """Format event list: time | subject | location | status.

    Example::

        2026-03-07 10:00-11:00 | Team standup | Room A | accepted
        2026-03-07 14:00-15:00 | 1:1 with Alice | Teams | tentativelyAccepted
    """
    if not events:
        return "No events found."

    lines: list[str] = []
    for event in events:
        parts: list[str] = []

        # Time range
        start = event.get("start", {}).get("dateTime", "")
        end = event.get("end", {}).get("dateTime", "")
        if event.get("isAllDay"):
            parts.append(f"{_compact_date(start)} (all day)")
        else:
            parts.append(f"{_compact_datetime(start)}-{_compact_time(end)}")

        # Subject
        subject = event.get("subject", "(no subject)")
        if len(subject) > 50:
            subject = subject[:47] + "..."
        parts.append(subject)

        # Location
        location = event.get("location", {}).get("displayName", "")
        if location:
            parts.append(location)

        # Status
        response = event.get("responseStatus", {}).get("response", "")
        if response and response != "none":
            parts.append(response)

        # Online meeting
        if event.get("isOnlineMeeting"):
            parts.append("online")

        # ID
        eid = event.get("id", "")
        if eid:
            parts.append(f"id={eid[:20]}...")

        lines.append(" | ".join(parts))

    return "\n".join(lines)


def format_event_detail(event: dict[str, Any]) -> str:
    """Format a single event with full detail."""
    lines: list[str] = []

    if subject := event.get("subject"):
        lines.append(f"Subject: {subject}")

    start = event.get("start", {})
    end = event.get("end", {})
    if event.get("isAllDay"):
        lines.append(f"When: {_compact_date(start.get('dateTime', ''))} (all day)")
    else:
        lines.append(f"Start: {start.get('dateTime', '?')} ({start.get('timeZone', '?')})")
        lines.append(f"End: {end.get('dateTime', '?')} ({end.get('timeZone', '?')})")

    location = event.get("location", {}).get("displayName", "")
    if location:
        lines.append(f"Location: {location}")

    organizer = event.get("organizer", {}).get("emailAddress", {})
    if organizer:
        lines.append(f"Organizer: {_format_email_address(organizer)}")

    attendees = event.get("attendees", [])
    if attendees:
        att_lines = []
        for att in attendees[:20]:
            addr = _format_email_address(att.get("emailAddress", {}))
            status = att.get("status", {}).get("response", "?")
            att_lines.append(f"  {addr} ({status})")
        lines.append(f"Attendees ({len(attendees)}):")
        lines.extend(att_lines)

    response = event.get("responseStatus", {}).get("response", "")
    if response:
        lines.append(f"Your response: {response}")

    if event.get("isOnlineMeeting"):
        url = event.get("onlineMeetingUrl", "")
        lines.append(f"Online meeting: {url}" if url else "Online meeting: yes")

    if eid := event.get("id"):
        lines.append(f"ID: {eid}")

    # Body
    body_data = event.get("body", {})
    content = body_data.get("content", "")
    if content:
        content_type = body_data.get("contentType", "text")
        if content_type.lower() == "html":
            content = _strip_html(content)
        if content.strip():
            lines.append("")
            lines.append(_truncate_body(content, 5000))

    return "\n".join(lines)


def format_freebusy(schedules: list[dict[str, Any]]) -> str:
    """Format free/busy schedules.

    Example::

        alice@example.com: BBBFFBBFF (B=busy, F=free)
          10:00-11:00 busy (Team standup)
          14:00-15:00 busy (1:1)
    """
    if not schedules:
        return "No schedule data."

    parts: list[str] = []
    for sched in schedules:
        email = sched.get("scheduleId", "?")
        availability = sched.get("availabilityView", "")
        parts.append(f"{email}: {availability}")

        items = sched.get("scheduleItems", [])
        for item in items:
            start = _compact_datetime(item.get("start", {}).get("dateTime", ""))
            end = _compact_time(item.get("end", {}).get("dateTime", ""))
            status = item.get("status", "?")
            subject = item.get("subject", "")
            line = f"  {start}-{end} {status}"
            if subject:
                line += f" ({subject})"
            parts.append(line)
        parts.append("")

    return "\n".join(parts).rstrip()


# ===========================================================================
# TASK FORMATTERS
# ===========================================================================


def format_task_lists(lists: list[dict[str, Any]]) -> str:
    """Format To Do task lists.

    Example::

        Tasks (owner) id=AAMk...
        Shopping (shared) id=AAMk...
    """
    if not lists:
        return "No task lists found."

    lines: list[str] = []
    for tl in lists:
        name = tl.get("displayName", "?")
        tags: list[str] = []
        if tl.get("isOwner"):
            tags.append("owner")
        if tl.get("isShared"):
            tags.append("shared")
        tag_str = f" ({', '.join(tags)})" if tags else ""
        tid = tl.get("id", "?")
        lines.append(f"{name}{tag_str} id={tid[:20]}...")
    return "\n".join(lines)


def format_task_list_items(tasks: list[dict[str, Any]]) -> str:
    """Format tasks: status | title | due | importance.

    Example::

        [ ] Buy groceries | due: 2026-03-08 | high | id=AAMk...
        [x] Send report | completed: 2026-03-06 | normal | id=AAMk...
    """
    if not tasks:
        return "No tasks found."

    lines: list[str] = []
    for task in tasks:
        parts: list[str] = []

        # Status checkbox
        status = task.get("status", "notStarted")
        checkbox = "[x]" if status == "completed" else "[ ]"
        parts.append(checkbox)

        # Title
        title = task.get("title", "(untitled)")
        parts.append(title)

        # Due date
        due = task.get("dueDateTime", {})
        if due:
            due_date = due.get("dateTime", "")
            if due_date:
                parts.append(f"due: {_compact_date(due_date)}")

        # Completed date
        completed = task.get("completedDateTime", {})
        if completed:
            comp_date = completed.get("dateTime", "")
            if comp_date:
                parts.append(f"completed: {_compact_date(comp_date)}")

        # Importance
        importance = task.get("importance", "normal")
        if importance != "normal":
            parts.append(importance)

        # ID
        tid = task.get("id", "")
        if tid:
            parts.append(f"id={tid[:20]}...")

        lines.append(" | ".join(parts))

    return "\n".join(lines)


def format_planner_plans(plans: list[dict[str, Any]]) -> str:
    """Format Planner plans."""
    if not plans:
        return "No Planner plans found."

    lines: list[str] = []
    for plan in plans:
        title = plan.get("title", "?")
        pid = plan.get("id", "?")
        lines.append(f"{title} id={pid[:20]}...")
    return "\n".join(lines)


def format_planner_tasks(tasks: list[dict[str, Any]]) -> str:
    """Format Planner tasks."""
    if not tasks:
        return "No Planner tasks found."

    lines: list[str] = []
    for task in tasks:
        parts: list[str] = []

        pct = task.get("percentComplete", 0)
        checkbox = "[x]" if pct == 100 else f"[{pct}%]" if pct > 0 else "[ ]"
        parts.append(checkbox)

        parts.append(task.get("title", "(untitled)"))

        due = task.get("dueDateTime", "")
        if due:
            parts.append(f"due: {_compact_date(due)}")

        priority = task.get("priority", 5)
        if priority <= 1:
            parts.append("urgent")
        elif priority <= 3:
            parts.append("important")

        tid = task.get("id", "")
        if tid:
            parts.append(f"id={tid[:20]}...")

        lines.append(" | ".join(parts))

    return "\n".join(lines)


# ===========================================================================
# META FORMATTERS
# ===========================================================================


def format_user_info(info: dict[str, Any]) -> str:
    """Format user profile info."""
    lines: list[str] = []
    if name := info.get("displayName"):
        lines.append(f"Name: {name}")
    if mail := info.get("mail"):
        lines.append(f"Email: {mail}")
    if upn := info.get("userPrincipalName"):
        lines.append(f"UPN: {upn}")
    return "\n".join(lines) if lines else str(info)


# ===========================================================================
# HELPERS
# ===========================================================================


def _compact_datetime(dt_str: str) -> str:
    """Convert ISO datetime to compact format: YYYY-MM-DD HH:MM."""
    if not dt_str:
        return "?"
    # Handle both 2026-03-07T10:30:00Z and 2026-03-07T10:30:00.0000000 formats
    return dt_str[:16].replace("T", " ")


def _compact_date(dt_str: str) -> str:
    """Extract just the date: YYYY-MM-DD."""
    if not dt_str:
        return "?"
    return dt_str[:10]


def _compact_time(dt_str: str) -> str:
    """Extract just the time: HH:MM."""
    if not dt_str:
        return "?"
    if "T" in dt_str:
        return dt_str.split("T")[1][:5]
    return dt_str[:5]


def _format_email_address(addr: dict[str, Any]) -> str:
    """Format an email address dict to 'Name <email>' or just 'email'."""
    name = addr.get("name", "")
    email = addr.get("address", "")
    if name and email:
        return f"{name} <{email}>"
    return email or name or "?"


def _strip_html(html: str) -> str:
    """Strip HTML tags and decode common entities. Simple and fast."""
    # Remove style/script blocks
    text = re.sub(r"<(style|script)[^>]*>.*?</\1>", "", html, flags=re.DOTALL | re.IGNORECASE)
    # Replace block elements with newlines
    text = re.sub(r"<(br|p|div|tr|li|h[1-6])[^>]*>", "\n", text, flags=re.IGNORECASE)
    # Remove all remaining tags
    text = re.sub(r"<[^>]+>", "", text)
    # Decode common entities
    text = text.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    text = text.replace("&quot;", '"').replace("&#39;", "'")
    # Collapse whitespace
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def _truncate_body(text: str, max_chars: int = MAX_BODY_CHARS) -> str:
    """Truncate long body text with annotation."""
    if len(text) <= max_chars:
        return text
    return text[:max_chars] + f"\n\n... (truncated, {len(text) - max_chars} more characters)"


def _human_size(size_bytes: int | float) -> str:
    """Convert bytes to human-readable size."""
    for unit in ("B", "KB", "MB", "GB"):
        if abs(size_bytes) < 1024:
            return f"{size_bytes:.1f} {unit}" if unit != "B" else f"{int(size_bytes)} B"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"
