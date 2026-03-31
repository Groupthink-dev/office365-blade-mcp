"""Microsoft Graph API client wrapper.

Wraps ``httpx`` with MSAL token management, typed exceptions, and convenience
methods for email, calendar, and tasks. All methods are synchronous — the server
wraps them with ``asyncio.to_thread()``.
"""

from __future__ import annotations

import logging
import re
from datetime import UTC, datetime
from typing import Any

import httpx

from office365_blade_mcp.auth import acquire_token
from office365_blade_mcp.models import (
    DEFAULT_LIMIT,
    EMAIL_LIST_FIELDS,
    EMAIL_READ_FIELDS,
    EVENT_LIST_FIELDS,
    EVENT_READ_FIELDS,
    GRAPH_BASE_URL,
    MAX_BATCH_SIZE,
    TASK_LIST_FIELDS,
    get_scopes,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------


class GraphError(Exception):
    """Base exception for Microsoft Graph API errors."""

    def __init__(self, message: str, status_code: int = 0, details: str = "") -> None:
        super().__init__(message)
        self.status_code = status_code
        self.details = details


class AuthError(GraphError):
    """Authentication or authorization failed."""


class NotFoundError(GraphError):
    """Requested resource not found."""


class RateLimitError(GraphError):
    """Rate limit exceeded — back off and retry."""


class ConnectionError(GraphError):  # noqa: A001
    """Cannot connect to Microsoft Graph API."""


class WriteDisabledError(GraphError):
    """Write operation attempted but O365_WRITE_ENABLED is not true."""


# ---------------------------------------------------------------------------
# Error classification
# ---------------------------------------------------------------------------


def _classify_error(status_code: int, message: str) -> GraphError:
    """Map HTTP status and message to a typed exception."""
    if status_code == 401 or status_code == 403:
        return AuthError(message, status_code)
    if status_code == 404:
        return NotFoundError(message, status_code)
    if status_code == 429:
        return RateLimitError(message, status_code)
    if status_code == 0:
        return ConnectionError(message, status_code)
    return GraphError(message, status_code)


def _scrub_token(text: str) -> str:
    """Remove access tokens from text to prevent leakage."""
    return re.sub(r"eyJ[a-zA-Z0-9_-]+\.eyJ[a-zA-Z0-9_-]+\.[a-zA-Z0-9_-]+", "[token]", text)


# ---------------------------------------------------------------------------
# Client
# ---------------------------------------------------------------------------


class GraphClient:
    """Microsoft Graph API client.

    Wraps ``httpx.Client`` with MSAL token management and typed exceptions.
    All methods are synchronous — the MCP server wraps them with
    ``asyncio.to_thread()`` to avoid blocking the event loop.
    """

    def __init__(self) -> None:
        self._http = httpx.Client(base_url=GRAPH_BASE_URL, timeout=30.0)
        self._access_token: str | None = None
        self._user_info: dict[str, Any] | None = None

    def _get_headers(self) -> dict[str, str]:
        """Get HTTP headers with a valid access token."""
        if self._access_token is None:
            scopes = get_scopes()
            self._access_token = acquire_token(scopes)
        return {
            "Authorization": f"Bearer {self._access_token}",
            "Content-Type": "application/json",
        }

    def _refresh_token(self) -> None:
        """Force token refresh on next request."""
        self._access_token = None

    def _request(self, method: str, url: str, **kwargs: Any) -> Any:
        """Make a Graph API request with error handling and token refresh."""
        try:
            resp = self._http.request(method, url, headers=self._get_headers(), **kwargs)
        except httpx.ConnectError as e:
            raise ConnectionError(f"Cannot connect to Graph API: {_scrub_token(str(e))}") from e
        except httpx.TimeoutException as e:
            raise ConnectionError(f"Graph API timeout: {_scrub_token(str(e))}") from e

        if resp.status_code == 401:
            # Token may have expired — refresh and retry once
            self._refresh_token()
            try:
                resp = self._http.request(method, url, headers=self._get_headers(), **kwargs)
            except httpx.ConnectError as e:
                raise ConnectionError(f"Cannot connect to Graph API: {_scrub_token(str(e))}") from e

        if resp.status_code >= 400:
            try:
                error_body = resp.json()
                error_msg = error_body.get("error", {}).get("message", resp.text[:500])
            except Exception:
                error_msg = resp.text[:500]
            raise _classify_error(resp.status_code, _scrub_token(error_msg))

        if resp.status_code == 204:
            return None
        return resp.json()

    def _get(self, url: str, params: dict[str, Any] | None = None) -> Any:
        return self._request("GET", url, params=params)

    def _post(self, url: str, json_data: dict[str, Any] | None = None) -> Any:
        return self._request("POST", url, json=json_data)

    def _patch(self, url: str, json_data: dict[str, Any]) -> Any:
        return self._request("PATCH", url, json=json_data)

    def _delete(self, url: str) -> Any:
        return self._request("DELETE", url)

    # ===================================================================
    # META
    # ===================================================================

    def get_user_info(self) -> dict[str, Any]:
        """Get current user profile."""
        if self._user_info is None:
            self._user_info = self._get("/me?$select=displayName,mail,userPrincipalName")
        return self._user_info

    # ===================================================================
    # EMAIL — READ
    # ===================================================================

    def get_mail_folders(self) -> list[dict[str, Any]]:
        """List all mail folders with counts."""
        data = self._get("/me/mailFolders?$top=100&$select=id,displayName,totalItemCount,unreadItemCount")
        return data.get("value", [])

    def search_emails(
        self,
        from_addr: str | None = None,
        to_addr: str | None = None,
        subject: str | None = None,
        body: str | None = None,
        after: str | None = None,
        before: str | None = None,
        folder_id: str | None = None,
        is_read: bool | None = None,
        importance: str | None = None,
        has_attachments: bool | None = None,
        limit: int = DEFAULT_LIMIT,
    ) -> tuple[list[dict[str, Any]], int]:
        """Search emails with filters. Returns (emails, total_count)."""
        filters: list[str] = []
        if from_addr:
            filters.append(f"from/emailAddress/address eq '{from_addr}'")
        if to_addr:
            filters.append(f"toRecipients/any(r:r/emailAddress/address eq '{to_addr}')")
        if subject:
            filters.append(f"contains(subject, '{_escape_odata(subject)}')")
        if body:
            filters.append(f"contains(body/content, '{_escape_odata(body)}')")
        if after:
            filters.append(f"receivedDateTime ge {after}T00:00:00Z")
        if before:
            filters.append(f"receivedDateTime lt {before}T23:59:59Z")
        if is_read is not None:
            filters.append(f"isRead eq {str(is_read).lower()}")
        if importance:
            filters.append(f"importance eq '{importance}'")
        if has_attachments is not None:
            filters.append(f"hasAttachments eq {str(has_attachments).lower()}")

        select = ",".join(EMAIL_LIST_FIELDS)
        base = f"/me/mailFolders/{folder_id}/messages" if folder_id else "/me/messages"
        params: dict[str, Any] = {
            "$select": select,
            "$top": min(limit, 50),
            "$orderby": "receivedDateTime desc",
            "$count": "true",
        }
        if filters:
            params["$filter"] = " and ".join(filters)

        data = self._get(base, params=params)
        messages = data.get("value", [])
        total = data.get("@odata.count", len(messages))
        return messages, total

    def get_email(self, message_id: str) -> dict[str, Any]:
        """Get a single email by ID with full content."""
        select = ",".join(EMAIL_READ_FIELDS)
        data = self._get(f"/me/messages/{message_id}?$select={select}")
        return data

    def get_email_thread(self, conversation_id: str, limit: int = DEFAULT_LIMIT) -> list[dict[str, Any]]:
        """Get all emails in a conversation thread."""
        select = ",".join(EMAIL_LIST_FIELDS + ["body"])
        params: dict[str, Any] = {
            "$filter": f"conversationId eq '{conversation_id}'",
            "$select": select,
            "$orderby": "receivedDateTime asc",
            "$top": min(limit, 50),
        }
        data = self._get("/me/messages", params=params)
        return data.get("value", [])

    def get_email_snippets(
        self,
        from_addr: str | None = None,
        subject: str | None = None,
        after: str | None = None,
        before: str | None = None,
        folder_id: str | None = None,
        limit: int = DEFAULT_LIMIT,
    ) -> tuple[list[dict[str, Any]], int]:
        """Search emails returning only preview snippets (token-efficient)."""
        select = "id,conversationId,receivedDateTime,subject,from,bodyPreview,isRead,flag"
        filters: list[str] = []
        if from_addr:
            filters.append(f"from/emailAddress/address eq '{from_addr}'")
        if subject:
            filters.append(f"contains(subject, '{_escape_odata(subject)}')")
        if after:
            filters.append(f"receivedDateTime ge {after}T00:00:00Z")
        if before:
            filters.append(f"receivedDateTime lt {before}T23:59:59Z")

        base = f"/me/mailFolders/{folder_id}/messages" if folder_id else "/me/messages"
        params: dict[str, Any] = {
            "$select": select,
            "$top": min(limit, 50),
            "$orderby": "receivedDateTime desc",
            "$count": "true",
        }
        if filters:
            params["$filter"] = " and ".join(filters)

        data = self._get(base, params=params)
        messages = data.get("value", [])
        total = data.get("@odata.count", len(messages))
        return messages, total

    def get_email_state(self) -> str:
        """Get a delta link for email change tracking."""
        data = self._get("/me/messages/delta?$select=id&$top=1")
        delta_link = data.get("@odata.deltaLink", "")
        if not delta_link:
            # First call returns nextLink, follow until deltaLink
            while "@odata.nextLink" in data:
                data = self._get(data["@odata.nextLink"].replace(GRAPH_BASE_URL, ""))
            delta_link = data.get("@odata.deltaLink", "")
        return delta_link

    def get_email_changes(self, delta_link: str) -> dict[str, Any]:
        """Get email changes since a delta link."""
        # Strip the base URL prefix if present
        url = delta_link.replace(GRAPH_BASE_URL, "")
        data = self._get(url)
        changes = data.get("value", [])
        new_delta = data.get("@odata.deltaLink", "")
        return {
            "changes": changes,
            "new_delta_link": new_delta,
            "count": len(changes),
        }

    def get_email_attachments(self, message_id: str) -> list[dict[str, Any]]:
        """List attachments for an email (metadata only, no content)."""
        data = self._get(f"/me/messages/{message_id}/attachments?$select=id,name,contentType,size")
        return data.get("value", [])

    # ===================================================================
    # EMAIL — WRITE
    # ===================================================================

    def send_email(
        self,
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
    ) -> str:
        """Send an email. Returns 'sent'."""
        message: dict[str, Any] = {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to],
        }
        if cc:
            message["ccRecipients"] = [{"emailAddress": {"address": addr}} for addr in cc]
        if bcc:
            message["bccRecipients"] = [{"emailAddress": {"address": addr}} for addr in bcc]

        self._post("/me/sendMail", json_data={"message": message, "saveToSentItems": True})
        return "sent"

    def reply_email(self, message_id: str, body: str, reply_all: bool = False) -> str:
        """Reply to an email. Returns 'replied'."""
        endpoint = f"/me/messages/{message_id}/replyAll" if reply_all else f"/me/messages/{message_id}/reply"
        self._post(endpoint, json_data={"comment": body})
        return "replied"

    def forward_email(self, message_id: str, to: list[str], comment: str = "") -> str:
        """Forward an email. Returns 'forwarded'."""
        self._post(
            f"/me/messages/{message_id}/forward",
            json_data={
                "comment": comment,
                "toRecipients": [{"emailAddress": {"address": addr}} for addr in to],
            },
        )
        return "forwarded"

    def flag_email(self, message_id: str, flag_status: str = "flagged") -> str:
        """Set flag on an email. Status: flagged, complete, notFlagged."""
        self._patch(f"/me/messages/{message_id}", json_data={"flag": {"flagStatus": flag_status}})
        return f"flag set to {flag_status}"

    def mark_email_read(self, message_id: str, is_read: bool = True) -> str:
        """Mark email as read or unread."""
        self._patch(f"/me/messages/{message_id}", json_data={"isRead": is_read})
        return "marked read" if is_read else "marked unread"

    def move_email(self, message_id: str, destination_folder_id: str) -> dict[str, Any]:
        """Move email to a folder. Returns moved message."""
        return self._post(
            f"/me/messages/{message_id}/move",
            json_data={"destinationId": destination_folder_id},
        )

    def delete_email(self, message_id: str) -> str:
        """Delete an email (moves to Deleted Items)."""
        self._delete(f"/me/messages/{message_id}")
        return "deleted"

    def bulk_email_action(
        self,
        ids: list[str],
        action: str,
        target_folder: str | None = None,
    ) -> int:
        """Perform bulk action on emails. Returns count affected."""
        capped = ids[: MAX_BATCH_SIZE]
        count = 0
        for msg_id in capped:
            try:
                if action == "mark_read":
                    self.mark_email_read(msg_id, is_read=True)
                elif action == "mark_unread":
                    self.mark_email_read(msg_id, is_read=False)
                elif action == "flag":
                    self.flag_email(msg_id, "flagged")
                elif action == "unflag":
                    self.flag_email(msg_id, "notFlagged")
                elif action == "move":
                    if not target_folder:
                        raise GraphError("target_folder required for move action")
                    self.move_email(msg_id, target_folder)
                elif action == "delete":
                    self.delete_email(msg_id)
                else:
                    raise GraphError(
                        f"Unknown action: {action}. "
                        "Valid: mark_read, mark_unread, flag, unflag, move, delete"
                    )
                count += 1
            except GraphError:
                logger.warning("Bulk action %s failed for %s", action, msg_id)
        return count

    # ===================================================================
    # CALENDAR — READ
    # ===================================================================

    def get_calendars(self) -> list[dict[str, Any]]:
        """List all calendars."""
        data = self._get("/me/calendars?$select=id,name,color,isDefaultCalendar,canEdit,owner")
        return data.get("value", [])

    def get_events(
        self,
        start: str,
        end: str,
        calendar_id: str | None = None,
        limit: int = DEFAULT_LIMIT,
    ) -> list[dict[str, Any]]:
        """Get events in a date range (calendarView)."""
        select = ",".join(EVENT_LIST_FIELDS)
        base = f"/me/calendars/{calendar_id}/calendarView" if calendar_id else "/me/calendarView"
        params: dict[str, Any] = {
            "startDateTime": f"{start}T00:00:00Z" if "T" not in start else start,
            "endDateTime": f"{end}T23:59:59Z" if "T" not in end else end,
            "$select": select,
            "$top": min(limit, 100),
            "$orderby": "start/dateTime asc",
        }
        data = self._get(base, params=params)
        return data.get("value", [])

    def get_event(self, event_id: str) -> dict[str, Any]:
        """Get a single event by ID with full detail."""
        select = ",".join(EVENT_READ_FIELDS)
        return self._get(f"/me/events/{event_id}?$select={select}")

    def search_events(
        self,
        query: str,
        start: str | None = None,
        end: str | None = None,
        limit: int = DEFAULT_LIMIT,
    ) -> list[dict[str, Any]]:
        """Search events by subject text."""
        select = ",".join(EVENT_LIST_FIELDS)
        filters = [f"contains(subject, '{_escape_odata(query)}')"]
        if start:
            filters.append(f"start/dateTime ge '{start}T00:00:00Z'")
        if end:
            filters.append(f"end/dateTime le '{end}T23:59:59Z'")
        params: dict[str, Any] = {
            "$filter": " and ".join(filters),
            "$select": select,
            "$top": min(limit, 50),
            "$orderby": "start/dateTime asc",
        }
        data = self._get("/me/events", params=params)
        return data.get("value", [])

    def get_freebusy(
        self,
        start: str,
        end: str,
        schedules: list[str] | None = None,
    ) -> list[dict[str, Any]]:
        """Get free/busy schedule for users."""
        user_info = self.get_user_info()
        if not schedules:
            email = user_info.get("mail") or user_info.get("userPrincipalName", "")
            schedules = [email]
        body = {
            "schedules": schedules,
            "startTime": {"dateTime": f"{start}T00:00:00", "timeZone": "UTC"},
            "endTime": {"dateTime": f"{end}T23:59:59", "timeZone": "UTC"},
            "availabilityViewInterval": 30,
        }
        data = self._post("/me/calendar/getSchedule", json_data=body)
        return data.get("value", [])

    # ===================================================================
    # CALENDAR — WRITE
    # ===================================================================

    def create_event(
        self,
        subject: str,
        start: str,
        end: str,
        body: str | None = None,
        location: str | None = None,
        attendees: list[str] | None = None,
        is_all_day: bool = False,
        calendar_id: str | None = None,
    ) -> dict[str, Any]:
        """Create a calendar event. Returns the created event."""
        event: dict[str, Any] = {
            "subject": subject,
            "start": {"dateTime": start, "timeZone": "UTC"},
            "end": {"dateTime": end, "timeZone": "UTC"},
            "isAllDay": is_all_day,
        }
        if body:
            event["body"] = {"contentType": "Text", "content": body}
        if location:
            event["location"] = {"displayName": location}
        if attendees:
            event["attendees"] = [
                {"emailAddress": {"address": addr}, "type": "required"} for addr in attendees
            ]

        base = f"/me/calendars/{calendar_id}/events" if calendar_id else "/me/events"
        return self._post(base, json_data=event)

    def update_event(self, event_id: str, updates: dict[str, Any]) -> dict[str, Any]:
        """Update a calendar event. Returns the updated event."""
        return self._patch(f"/me/events/{event_id}", json_data=updates)

    def delete_event(self, event_id: str) -> str:
        """Delete a calendar event."""
        self._delete(f"/me/events/{event_id}")
        return "deleted"

    def respond_event(self, event_id: str, response: str, comment: str = "") -> str:
        """Respond to an event invitation: accept, decline, tentativelyAccept."""
        valid = ("accept", "decline", "tentativelyAccept")
        if response not in valid:
            raise GraphError(f"Invalid response: {response}. Valid: {', '.join(valid)}")
        self._post(f"/me/events/{event_id}/{response}", json_data={"comment": comment, "sendResponse": True})
        return f"responded: {response}"

    # ===================================================================
    # TASKS (To Do) — READ
    # ===================================================================

    def get_task_lists(self) -> list[dict[str, Any]]:
        """List all To Do task lists."""
        data = self._get("/me/todo/lists?$select=id,displayName,isOwner,isShared")
        return data.get("value", [])

    def get_tasks(
        self,
        list_id: str,
        status: str | None = None,
        limit: int = DEFAULT_LIMIT,
    ) -> list[dict[str, Any]]:
        """Get tasks from a task list."""
        select = ",".join(TASK_LIST_FIELDS)
        params: dict[str, Any] = {
            "$select": select,
            "$top": min(limit, 100),
            "$orderby": "createdDateTime desc",
        }
        if status:
            params["$filter"] = f"status eq '{status}'"
        data = self._get(f"/me/todo/lists/{list_id}/tasks", params=params)
        return data.get("value", [])

    def get_task(self, list_id: str, task_id: str) -> dict[str, Any]:
        """Get a single task by ID."""
        select = ",".join(TASK_LIST_FIELDS)
        return self._get(f"/me/todo/lists/{list_id}/tasks/{task_id}?$select={select}")

    def search_tasks(self, list_id: str, query: str, limit: int = DEFAULT_LIMIT) -> list[dict[str, Any]]:
        """Search tasks by title text."""
        select = ",".join(TASK_LIST_FIELDS)
        params: dict[str, Any] = {
            "$filter": f"contains(title, '{_escape_odata(query)}')",
            "$select": select,
            "$top": min(limit, 50),
        }
        data = self._get(f"/me/todo/lists/{list_id}/tasks", params=params)
        return data.get("value", [])

    # ===================================================================
    # TASKS (To Do) — WRITE
    # ===================================================================

    def create_task(
        self,
        list_id: str,
        title: str,
        body: str | None = None,
        due_date: str | None = None,
        importance: str = "normal",
    ) -> dict[str, Any]:
        """Create a task in a list. Returns the created task."""
        task: dict[str, Any] = {
            "title": title,
            "importance": importance,
        }
        if body:
            task["body"] = {"contentType": "text", "content": body}
        if due_date:
            task["dueDateTime"] = {"dateTime": f"{due_date}T00:00:00", "timeZone": "UTC"}
        return self._post(f"/me/todo/lists/{list_id}/tasks", json_data=task)

    def update_task(self, list_id: str, task_id: str, updates: dict[str, Any]) -> dict[str, Any]:
        """Update a task. Returns the updated task."""
        return self._patch(f"/me/todo/lists/{list_id}/tasks/{task_id}", json_data=updates)

    def complete_task(self, list_id: str, task_id: str) -> dict[str, Any]:
        """Mark a task as completed."""
        now = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
        return self._patch(
            f"/me/todo/lists/{list_id}/tasks/{task_id}",
            json_data={
                "status": "completed",
                "completedDateTime": {"dateTime": now, "timeZone": "UTC"},
            },
        )

    def delete_task(self, list_id: str, task_id: str) -> str:
        """Delete a task."""
        self._delete(f"/me/todo/lists/{list_id}/tasks/{task_id}")
        return "deleted"

    # ===================================================================
    # PLANNER — READ
    # ===================================================================

    def get_planner_plans(self) -> list[dict[str, Any]]:
        """List Planner plans the user has access to."""
        data = self._get("/me/planner/plans?$select=id,title,createdDateTime,owner")
        return data.get("value", [])

    def get_planner_tasks(self, plan_id: str, limit: int = DEFAULT_LIMIT) -> list[dict[str, Any]]:
        """Get tasks from a Planner plan."""
        params: dict[str, Any] = {"$top": min(limit, 100)}
        data = self._get(f"/planner/plans/{plan_id}/tasks", params=params)
        return data.get("value", [])

    # ===================================================================
    # PLANNER — WRITE
    # ===================================================================

    def create_planner_task(
        self,
        plan_id: str,
        bucket_id: str,
        title: str,
        assigned_to: str | None = None,
    ) -> dict[str, Any]:
        """Create a Planner task."""
        task: dict[str, Any] = {
            "planId": plan_id,
            "bucketId": bucket_id,
            "title": title,
        }
        if assigned_to:
            task["assignments"] = {
                assigned_to: {"@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !"}
            }
        return self._post("/planner/tasks", json_data=task)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _escape_odata(value: str) -> str:
    """Escape a string for use in OData filter expressions."""
    return value.replace("'", "''")
