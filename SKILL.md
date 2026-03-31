---
name: office365-blade
description: Microsoft 365 operations via Graph API — email, calendar, tasks
version: 0.1.0
permissions:
  read:
    - o365_info
    - email_folders
    - email_search
    - email_read
    - email_thread
    - email_snippets
    - email_attachments
    - email_state
    - email_changes
    - cal_calendars
    - cal_events
    - cal_event
    - cal_search
    - cal_today
    - cal_week
    - cal_freebusy
    - cal_batch
    - task_lists
    - task_list
    - task_search
    - task_today
    - task_inbox
    - planner_plans
    - planner_tasks
  write:
    - email_send
    - email_reply
    - email_forward
    - email_flag
    - email_move
    - email_delete
    - email_bulk
    - cal_respond
    - cal_create
    - cal_update
    - cal_delete
    - task_create
    - task_update
    - task_complete
---

# Office 365 Blade MCP — Skill Guide

## Token Efficiency Rules (MANDATORY)

1. **Use `email_snippets` before `email_read`** — preview mode is 90% fewer tokens
2. **Use `limit=` on all search/list tools** — default is 20, reduce for browsing
3. **Use `email_folders` to get IDs** — folder IDs are required for scoped searches
4. **Use `cal_today` / `cal_week` for common views** — avoids manual date calculation
5. **Use `cal_freebusy` for availability** — returns only busy periods, not full events
6. **Use `task_inbox` for quick task overview** — auto-finds default list
7. **Use `email_thread` for conversations** — more efficient than reading individual emails
8. **Never read all emails** — always filter by date, sender, folder, or flags

## Quick Start — 5 Most Common Operations

```
email_search limit=10                              → Recent emails
email_read id="AAMk..."                            → Read specific email
cal_today                                          → Today's events
task_inbox                                         → Pending tasks
o365_info                                          → Health check
```

## Tool Reference

### Meta
- **o365_info** — Account name, email, UPN. Health check.

### Email Read
- **email_folders** — All mail folders with ID, name, total/unread.
- **email_search** — Search with filters (from, to, subject, date, folder, read status).
- **email_read** — Full email: headers + body.
- **email_thread** — Chronological conversation view by conversation ID.
- **email_snippets** — Preview snippets only (token-efficient). Best for browsing.
- **email_attachments** — Attachment metadata (name, type, size). No content download.

### Email State & Changes
- **email_state** — Get delta link for change tracking.
- **email_changes** — Incremental changes since a delta link. Created/removed IDs.

### Email Write (requires O365_WRITE_ENABLED=true)
- **email_send** — Send new email. Comma-separated recipients.
- **email_reply** — Reply (or reply all) preserving threading.
- **email_forward** — Forward with optional comment.
- **email_flag** — Flag/unflag or mark read/unread.
- **email_move** — Move to folder by ID.
- **email_delete** — Delete (moves to Deleted Items). Requires `confirm=true`.
- **email_bulk** — Bulk action on up to 50 emails.

### Calendar Read
- **cal_calendars** — All calendars with ID, name, permissions.
- **cal_events** — Events in a date range from one calendar.
- **cal_event** — Full event detail: attendees, body, recurrence.
- **cal_search** — Search events by subject text.
- **cal_today** — Today's events (convenience shortcut).
- **cal_week** — This week's events Mon-Sun (convenience shortcut).
- **cal_freebusy** — Free/busy availability for one or more users.
- **cal_batch** — Events from multiple calendars in one call.

### Calendar Write (requires O365_WRITE_ENABLED=true)
- **cal_respond** — Accept/decline/tentative on event invitations.
- **cal_create** — Create event with attendees, location, body.
- **cal_update** — Update event fields.
- **cal_delete** — Delete event. Requires `confirm=true`.

### Tasks (To Do) Read
- **task_lists** — All To Do lists with ID and ownership.
- **task_list** — Tasks from a specific list. Filter by status.
- **task_search** — Search tasks by title.
- **task_today** — Tasks due today from a list.
- **task_inbox** — All pending tasks from default list.

### Tasks Write (requires O365_WRITE_ENABLED=true)
- **task_create** — Create task with title, body, due date, importance.
- **task_update** — Update task fields.
- **task_complete** — Mark task as completed.

### Planner
- **planner_plans** — List accessible Planner plans.
- **planner_tasks** — Tasks from a Planner plan.

## Workflow Examples

### Email Triage
```
1. email_folders                                → Get Inbox folder ID
2. email_snippets folder_id="..." limit=20      → Preview recent emails
3. email_read id="AAMk..."                      → Read specific email
4. email_thread conversation_id="..."           → Get full conversation
5. email_flag ids="..." action="mark_read"      → Mark as read
```

### Schedule a Meeting
```
1. cal_freebusy start="2026-03-08" end="2026-03-08" schedules="alice@example.com,bob@example.com"
2. cal_create subject="Team sync" start="2026-03-08T14:00:00" end="2026-03-08T15:00:00" attendees="alice@example.com,bob@example.com" location="Room A"
```

### Task Management
```
1. task_lists                                   → Find task list IDs
2. task_list list_id="..." status="notStarted"  → View pending tasks
3. task_create list_id="..." title="Review PR" due_date="2026-03-08" importance="high"
4. task_complete list_id="..." task_id="..."     → Mark done
```

### Incremental Email Sync
```
1. email_state                                  → Get delta link
   (persist delta link for next run)
2. email_changes delta_link="..."               → Get changed message IDs
3. email_read id="..."                          → Read changed messages
```

## Common Parameters

| Parameter | Description | Example |
|-----------|-------------|---------|
| `id` | Resource ID | `id="AAMk..."` |
| `ids` | Comma-separated IDs | `ids="AAMk...,AAMk..."` |
| `from_addr` | Sender email filter | `from_addr="alice@example.com"` |
| `subject` | Subject text filter | `subject="Meeting"` |
| `after` / `before` | Date range (YYYY-MM-DD) | `after="2026-03-01"` |
| `folder_id` | Mail folder ID | `folder_id="AAMk..."` |
| `calendar_id` | Calendar ID | `calendar_id="AAMk..."` |
| `list_id` | Task list ID | `list_id="AAMk..."` |
| `limit` | Max results (default: 20) | `limit=10` |
| `confirm` | Safety gate for destructive ops | `confirm=true` |

## Security Notes

- Write operations blocked unless `O365_WRITE_ENABLED=true`
- `email_delete` and `cal_delete` require `confirm=true`
- `email_bulk` capped at 50 emails per operation
- Access tokens never appear in tool output
- Token cache file has 0600 permissions (owner-only)
- Supports device_code (interactive) and client_credentials (headless) auth
